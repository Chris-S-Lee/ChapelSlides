import React, { useState } from 'react';
import pptxgen from "pptxgenjs";

// ─────────────────────────────────────────────────────────────────
// 1. 스타일 및 좌표 설정 (외곽선 및 투명도 100% 추가)
// ─────────────────────────────────────────────────────────────────
// 투명도 100% 외곽선 설정 (나중에 사용자가 PPT에서 색만 채우면 나타남)
const GHOST_OUTLINE = { color: "000000", size: 0 };

const STYLE_CONFIG = {
  TITLE_KOR: { 
    x: 2.2302, y: 0.875, w: 17.7143, h: 1.376, 
    fontFace: "페이퍼로지 7 Bold", fontSize: 84, color: "FFFFFF", 
    align: "right", valign: "bottom",
    // outline: GHOST_OUTLINE 
  },
  TITLE_ENG: { 
    x: 2.2292, y: 2.4574, w: 17.7708, h: 1.0962, 
    fontFace: "페이퍼로지 5 Medium", fontSize: 54, color: "FFFFFF", 
    align: "right", valign: "top",
    // outline: GHOST_OUTLINE
  },
  LYRICS_KOR: { 
    x: 0, y: 0.1816, w: 20, h: 2.235, 
    fontFace: "페이퍼로지 5 Medium", fontSize: 72, color: "FFFFFF", 
    align: "center", valign: "top", lineSpacingMultiple: 0.95,
    // outline: GHOST_OUTLINE
  },
  LYRICS_ENG: { 
    x: 0, y: 2.5322, w: 20, h: 1.9295, 
    fontFace: "페이퍼로지 4 Regular", fontSize: 40, color: "FFFFFF", 
    align: "center", valign: "top", lineSpacingMultiple: 0.90,
    // outline: GHOST_OUTLINE
  },
};

// ─────────────────────────────────────────────────────────────────
// 2. 유틸리티 함수
// ─────────────────────────────────────────────────────────────────
function parseLyrics(raw) {
  const lines = raw.split('\n').map(l => l.trim());
  const sections = [];
  let name = "Intro", buf = [];
  lines.forEach(line => {
    const tag = line.match(/^\[(.+)\]$/);
    if (tag) {
      if (buf.length) sections.push({ name, lines: [...buf] });
      name = tag[1]; buf = [];
    } else if (line) buf.push(line);
  });
  if (buf.length) sections.push({ name, lines: [...buf] });
  return sections;
}

const isKor = t => /[ㄱ-ㅎㅏ-ㅣ가-힣]/.test(t);

// ─────────────────────────────────────────────────────────────────
// 3. 슬라이드 마스터 정의
// ─────────────────────────────────────────────────────────────────
function defineMasters(pptx) {
  const lyBg = "/template_bg.png"; 
  const tiBg = "/template_title_bg.png";

  pptx.defineSlideMaster({
    title: "LYRICS_MASTER",
    objects: [
      { image: { x: 0, y: 0, w: "100%", h: "100%", path: lyBg } },
      { placeholder: { options: { name: "korText", type: "body", ...STYLE_CONFIG.LYRICS_KOR }, text: "" } },
      { placeholder: { options: { name: "engText", type: "body", ...STYLE_CONFIG.LYRICS_ENG }, text: "" } },
    ],
  });

  pptx.defineSlideMaster({
    title: "TITLE_MASTER",
    objects: [
      { image: { x: 0, y: 0, w: "100%", h: "100%", path: tiBg } },
      { placeholder: { options: { name: "songTitleKor", type: "title", ...STYLE_CONFIG.TITLE_KOR }, text: "" } },
      { placeholder: { options: { name: "songTitleEng", type: "body", ...STYLE_CONFIG.TITLE_ENG }, text: "" } },
    ],
  });
}

// ─────────────────────────────────────────────────────────────────
// 4. PPT 생성 로직
// ─────────────────────────────────────────────────────────────────
function generatePPT(title, rawLyrics) {
  if (!rawLyrics.trim()) return alert("가사를 입력해주세요.");

  const pptx = new pptxgen();
  pptx.defineLayout({ name: "W20H11", width: 20, height: 11.25 });
  pptx.layout = "W20H11";
  
  defineMasters(pptx);
  const sections = parseLyrics(rawLyrics);

  sections.forEach((sec, idx) => {
    pptx.addSection({ title: sec.name });

    if (idx === 0) {
      const slide = pptx.addSlide({ masterName: "TITLE_MASTER", sectionTitle: sec.name });
      const kor = sec.lines.find(isKor) || title || sec.name;
      const eng = sec.lines.find(l => !isKor(l)) || "";

      slide.addText(kor, { placeholder: "songTitleKor", ...STYLE_CONFIG.TITLE_KOR });
      if (eng) slide.addText(eng, { placeholder: "songTitleEng", ...STYLE_CONFIG.TITLE_ENG });
    } else {
      const korLines = sec.lines.filter(isKor);
      const engLines = sec.lines.filter(l => !isKor(l));
      const maxLines = Math.max(korLines.length, engLines.length);

      for (let i = 0; i < maxLines; i += 2) {
        const slide = pptx.addSlide({ masterName: "LYRICS_MASTER", sectionTitle: sec.name });
        const korPart = korLines.slice(i, i + 2).join('\n');
        const engPart = engLines.slice(i, i + 2).join('\n');
        
        if (korPart) slide.addText(korPart, { placeholder: "korText", ...STYLE_CONFIG.LYRICS_KOR });
        if (engPart) slide.addText(engPart, { placeholder: "engText", ...STYLE_CONFIG.LYRICS_ENG });
      }
    }
  });

  pptx.writeFile({ fileName: `${title || "Worship"}.pptx` });
}

// ─────────────────────────────────────────────────────────────────
// 5. UI (React App)
// ─────────────────────────────────────────────────────────────────
export default function App() {
  const [title, setTitle] = useState('');
  const [rawLyrics, setRawLyrics] = useState('');

  const handleFile = e => {
    const f = e.target.files[0];
    if (!f) return;
    setTitle(f.name.replace(/\.[^/.]+$/, ""));
    const r = new FileReader();
    r.onload = ev => setRawLyrics(ev.target.result);
    r.readAsText(f, "UTF-8");
  };

  const inputStyle = {
    width: "100%", padding: "13px 16px", boxSizing: "border-box", borderRadius: "10px",
    background: "#f2f2f2", border: "1px dashed #0a0a0a", outline: "none", fontSize: "14px"
  };

  return (
    <div style={{ minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "sans-serif", backgroundColor: "#fafafa" }}>
      <div style={{ width: "100%", maxWidth: "680px", padding: "40px", backgroundColor: "#fff", border: "1px solid #ddd", borderRadius: "20px", boxShadow: "0 10px 25px rgba(0,0,0,0.05)" }}>
        <h1 style={{ textAlign: "center", fontSize: "24px", color: "#333", marginBottom: "30px" }}>Worship PPT Generator</h1>
        
        <label style={{ display: "block", background: "#f8f9fa", padding: "20px", borderRadius: "10px", cursor: "pointer", marginBottom: "15px", border: "1px dashed #bbb", textAlign: "center" }}>
          <strong>📤 TXT 파일 업로드</strong>
          <input type="file" accept=".txt" onChange={handleFile} style={{ display: "none" }} />
        </label>

        <input type="text" placeholder="찬양 제목" value={title} onChange={e => setTitle(e.target.value)} style={{ ...inputStyle, marginBottom: "10px" }} />
        
        <textarea
          placeholder="[섹션명]을 포함한 가사를 입력하세요."
          value={rawLyrics}
          onChange={e => setRawLyrics(e.target.value)}
          style={{ ...inputStyle, height: "300px", resize: "vertical", marginBottom: "20px", lineHeight: "1.6" }}
        />

        <button
          onClick={() => generatePPT(title, rawLyrics)}
          style={{ width: "100%", padding: "16px", background: "#007bff", color: "#fff", border: "none", borderRadius: "10px", fontWeight: "bold", cursor: "pointer", fontSize: "16px" }}
        >
          PPT 생성 및 다운로드
        </button>
      </div>
    </div>
  );
}