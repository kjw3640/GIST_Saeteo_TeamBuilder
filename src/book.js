// Read only the first worksheet, shared strings, and workbook metadata.
// Never visit relationship URLs or evaluate formulas from a workbook.
const UTF8 = new TextDecoder("utf-8", {fatal:true});
const enc = new TextEncoder();
const MB = 1024 * 1024;
const sheetNS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const relNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
function fail(message) { throw new Error(message); }
function descendants(root, tag) { return Array.from(root.getElementsByTagNameNS("*", tag)); }
function columnIndex(letters) {
  let n = 0;
  for (const c of letters) n = n * 26 + c.charCodeAt(0) - 64;
  return n - 1;
}
function zipDirectory(bytes) {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const u16 = i => view.getUint16(i,true), u32 = i => view.getUint32(i,true);
  const bad = () => fail("표준 .xlsx 통합 문서를 선택하세요. 암호화되거나 손상된 파일은 읽을 수 없습니다.");
  let end = bytes.length - 22;
  while (end >= Math.max(0, bytes.length-65557) && (u32(end)!==0x06054b50 || end+22+u16(end+20)!==bytes.length)) end--;
  if (end < 0 || u32(end)!==0x06054b50 || u16(end+4) || u16(end+6)) bad();
  const count = u16(end+10), entries = new Map();
  if (count > 2048) fail("파일 안의 항목이 너무 많습니다.");
  let at = u32(end+16);
  for (let i=0;i<count;i++) {
    if (at+46>end || u32(at)!==0x02014b50) bad();
    const flags=u16(at+8), method=u16(at+10), packed=u32(at+20), size=u32(at+24);
    const nameLen=u16(at+28), next=at+46+nameLen+u16(at+30)+u16(at+32);
    if (next>end || flags&1) bad();
    const name=UTF8.decode(bytes.subarray(at+46,at+46+nameLen)), local=u32(at+42);
    if (local+30>end || u32(local)!==0x04034b50 || entries.has(name)) bad();
    const offset=local+30+u16(local+26)+u16(local+28);
    if (offset+packed>end) bad();
    entries.set(name,{offset,packed,size,method});
    at=next;
  }
  let budget=24*MB;
  return async path => {
    const e=entries.get(path);
    if (!e) return null;
    if (e.size>budget || e.size>12*MB || ![0,8].includes(e.method)) fail("파일 압축 형식 또는 크기 제한을 확인하세요.");
    const data=bytes.subarray(e.offset,e.offset+e.packed);
    let stream=new Blob([data]).stream();
    if (e.method===8) stream=stream.pipeThrough(new DecompressionStream("deflate-raw"));
    const reader=stream.getReader(), parts=[];
    let used=0;
    while(true) {
      const {done,value}=await reader.read();
      if(done) break;
      used+=value.length;
      if(used>e.size || used>budget) { await reader.cancel(); fail("엑셀 압축 데이터 크기가 잘못되었습니다."); }
      parts.push(value);
    }
    if(used!==e.size) bad();
    budget-=used;
    const xml=UTF8.decode(await new Blob(parts).arrayBuffer());
    if (/<!DOCTYPE|<!ENTITY/i.test(xml)) fail("지원하지 않는 XML 선언이 있습니다.");
    const doc=new DOMParser().parseFromString(xml,"application/xml");
    if(descendants(doc,"parsererror").length) bad();
    return doc;
  };
}
async function readBook(file) {
  if (!/\.xlsx$/i.test(file.name) || file.size>8*MB) fail("8 MB 이하의 .xlsx 파일을 선택하세요.");
  const xml=zipDirectory(new Uint8Array(await file.arrayBuffer()));
  const book=await xml("xl/workbook.xml"), rel=await xml("xl/_rels/workbook.xml.rels");
  if(!book || !rel) fail("엑셀 통합 문서 정보를 찾을 수 없습니다.");
  const first=descendants(book,"sheet")[0], key=first?.getAttributeNS(relNS,"id");
  const rels=descendants(rel,"Relationship");
  function part(r) {
    const target=r?.getAttribute("Target") || "";
    if (!target || r.getAttribute("TargetMode")==="External" || /[:\\?#]|(?:^|\/)\.\.(?:\/|$)/.test(target)) fail("지원하지 않는 엑셀 시트 연결입니다.");
    return target.startsWith("/") ? target.slice(1) : "xl/"+target.replace(/^\.\//,"");
  }
  const sharedRel=rels.find(r=>r.getAttribute("Type")?.endsWith("/sharedStrings"));
  const sharedDoc=sharedRel ? await xml(part(sharedRel)) : null;
  const inlineText=el=>descendants(el,"t").map(t=>t.textContent).join("");
  const strings=sharedDoc ? descendants(sharedDoc,"si").map(inlineText) : [];
  const data=await xml(part(rels.find(r=>r.getAttribute("Id")===key)));
  if(!data) fail("첫 번째 시트를 읽을 수 없습니다.");
  const rows=[];
  const sheetData=descendants(data,"sheetData")[0];
  if(!sheetData) fail("빈 시트입니다.");
  for(const row of sheetData.children) {
    const cells=[];
    for(const c of row.children) {
      if(c.localName!=="c") continue;
      const ref=/^([A-Z]{1,3})\d+$/.exec(c.getAttribute("r")||"");
      if(!ref) continue;
      const col=columnIndex(ref[1]);
      if(col>=128) continue;
      const type=c.getAttribute("t"), v=descendants(c,"v")[0]?.textContent||"";
      const value=type==="inlineStr" ? inlineText(c) : type==="s" ? strings[Number(v)]||"" : v;
      if(value.length>32767) fail("셀 내용이 너무 깁니다.");
      if(value.trim()) cells[col]=value.trim();
    }
    if(cells.length) rows.push({number:Number(row.getAttribute("r")),cells});
    if(rows.length>2501) fail("파일당 최대 2,500행까지 읽을 수 있습니다.");
  }
  if(rows.length<2) fail("첫 번째 시트에 열 제목과 명단을 넣어주세요.");
  return rows;
}

// Export a fresh OOXML package containing only explicitly selected columns.
// All cell values use inline strings, including text beginning with '='.
const escapeXML=s=>String(s??"").replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g,"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");
const crcTable=Array.from({length:256},(_,i)=>{let c=i; for(let j=0;j<8;j++) c=c&1?0xedb88320^(c>>>1):c>>>1; return c>>>0;});
function zipStore(files) {
  const chunks=[], central=[];
  let offset=0, centralSize=0;
  for(const [path,xml] of files) {
    const name=enc.encode(path), data=enc.encode(xml);
    let crc=0xffffffff; for(const b of data) crc=(crc>>>8)^crcTable[(crc^b)&255]; crc=(crc^0xffffffff)>>>0;
    function header(dir) {
      const base=dir?46:30, b=new Uint8Array(base+name.length), v=new DataView(b.buffer);
      const w=(pos,n)=>v.setUint16(pos,n,true), d=(pos,n)=>v.setUint32(pos,n,true);
      d(0,dir?0x02014b50:0x04034b50);
      const diff=dir?2:0;
      if(dir) { w(4,20); d(42,offset); }
      w(4+diff,20); w(12+diff,33); d(14+diff,crc); d(18+diff,data.length); d(22+diff,data.length); w(26+diff,name.length);
      b.set(name,base); return b;
    }
    const local=header(false), cd=header(true);
    chunks.push(local,data);central.push(cd);centralSize+=cd.length;offset+=local.length+data.length;
  }
  const tail=new Uint8Array(22), v=new DataView(tail.buffer);
  v.setUint32(0,0x06054b50,true); v.setUint16(8,files.length,true); v.setUint16(10,files.length,true);
  v.setUint32(12,centralSize,true); v.setUint32(16,offset,true);
  return new Blob([...chunks,...central,tail],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
}
function makeWorkbook(sheets) {
  const pkg="http://schemas.openxmlformats.org/package/2006/relationships";
  const mime="application/vnd.openxmlformats-officedocument.spreadsheetml.";
  const relation=(id,kind,path)=>`<Relationship Id="r${id}" Type="${relNS}/${kind}" Target="${path}"/>`;
  const parts=[], names=[], links=[], types=[];
  sheets.forEach(({name,rows,widths,headingRows=[0]},i)=>{
    const n=i+1, path="worksheets/sheet"+n+".xml";
    names.push(`<sheet name="${escapeXML(name)}" sheetId="${n}" r:id="r${n}"/>`);
    links.push(relation(n,"worksheet",path));
    types.push(`<Override PartName="/xl/${path}" ContentType="${mime}worksheet+xml"/>`);
    const cells=rows.map((row,r)=>`<row r="${r+1}" ht="21" customHeight="1">${row.map((v,c)=>`<c r="${String.fromCharCode(65+c)}${r+1}" t="inlineStr" s="${headingRows.includes(r)?1:0}"><is><t xml:space="preserve">${escapeXML(v)}</t></is></c>`).join("")}</row>`).join("");
    parts.push(["xl/"+path,`<worksheet xmlns="${sheetNS}"><sheetViews><sheetView workbookViewId="0" showGridLines="0"><pane ySplit="1" topLeftCell="A2" state="frozen"/></sheetView></sheetViews><cols>${widths.map((w,i)=>`<col min="${i+1}" max="${i+1}" width="${w}" customWidth="1"/>`).join("")}</cols><sheetData>${cells}</sheetData></worksheet>`]);
  });
  const styles=`<styleSheet xmlns="${sheetNS}"><fonts count="2"><font><sz val="11"/><name val="맑은 고딕"/></font><font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="맑은 고딕"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF245A48"/></patternFill></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="2"><xf xfId="0"/><xf xfId="0" fontId="1" fillId="2" applyFont="1" applyFill="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;
  parts.push(["xl/styles.xml",styles],["xl/workbook.xml",`<workbook xmlns="${sheetNS}" xmlns:r="${relNS}"><sheets>${names.join("")}</sheets></workbook>`],["xl/_rels/workbook.xml.rels",`<Relationships xmlns="${pkg}">${links.join("")}${relation("style","styles","styles.xml")}</Relationships>`],["_rels/.rels",`<Relationships xmlns="${pkg}">${relation(1,"officeDocument","xl/workbook.xml")}</Relationships>`],["[Content_Types].xml",`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="${mime}sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="${mime}styles+xml"/>${types.join("")}</Types>`]);
  return zipStore(parts);
}
