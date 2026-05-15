import React, { useState } from 'react';
import pptxgen from "pptxgenjs";

// ─────────────────────────────────────────────────────────────────
//  Template.pptx에서 XML 직접 추출한 정확한 값
// ─────────────────────────────────────────────────────────────────

// 배경 그라데이션
// ang="16200038" (OOXML 1/60000°) = 270° → 아래(검정)에서 위(녹색)로
// pos=0: #000000, pos=75%(75200): #223415, pos=100%(100000): #223415
const BG_GRADIENT_STOPS = [
  { position:   0, color: "000000" },
  { position:  75, color: "223415" },
  { position: 100, color: "223415" },
];

// 가사 슬라이드 placeholder 위치 (slideLayout2.xml, EMU→인치 변환)
const LYRICS_KOR = {
  x: 0, y: 0.1816, w: 20, h: 2.235,
  font: "페이퍼로지 5 Medium",
  size: 72,          // sz="7200"
  lineSpacing: 95,   // spcPct val="95000"
  align: "center", color: "FFFFFF",
};
const LYRICS_ENG = {
  x: 0, y: 2.5322, w: 20, h: 1.9295,
  font: "페이퍼로지 4 Regular",
  size: 40,          // sz="4000"
  lineSpacing: 90,   // spcPct val="90000"
  align: "center", color: "FFFFFF",
};

// 제목 슬라이드 위치 (slideLayout1.xml)
const TITLE_KOR = {
  x: 2.2302, y: 0.875, w: 17.7143, h: 1.376,
  font: "페이퍼로지 7 Bold",
  size: 84,          // sz="8400"
  lineSpacing: 90,
  align: "right", color: "FFFFFF",
};
const TITLE_ENG = {
  x: 2.2292, y: 2.4574, w: 17.7708, h: 1.0962,
  font: "페이퍼로지 5 Medium",
  size: 54,          // sz="5400"
  lineSpacing: 90,
  align: "right", color: "FFFFFF",
};
// 구분선: x=6961239EMU, y=2058326EMU, cx=11326761EMU, thickness=76200EMU(≈6pt)
const TITLE_LINE = { x: 7.6129, y: 2.251, w: 12.3871, thickness: 6 };

// ─────────────────────────────────────────────────────────────────
//  슬라이드 마스터 정의
// ─────────────────────────────────────────────────────────────────
function defineMasters(pptx) {
  // pptxgenjs gradFill: angle=270 → 아래→위 (하단 검정, 상단 녹색)
  const bgObj = {
    rect: {
      x: 0, y: 0, w: "100%", h: "100%",
      fill: {
        type: "gradient",
        angle: 270,
        stops: BG_GRADIENT_STOPS,
      },
    },
  };

  // ── 마스터 1: 가사 슬라이드 ──
  pptx.defineSlideMaster({
    title: "LYRICS_MASTER",
    background: { color: "000000" },
    objects: [
      bgObj,
      {
        placeholder: {
          options: {
            name: "korText", type: "body",
            x: LYRICS_KOR.x, y: LYRICS_KOR.y,
            w: LYRICS_KOR.w, h: LYRICS_KOR.h,
            fontSize: LYRICS_KOR.size,
            fontFace: LYRICS_KOR.font,
            color: LYRICS_KOR.color,
            align: LYRICS_KOR.align,
            valign: "top",
            lineSpacingMultiple: LYRICS_KOR.lineSpacing / 100,
            paraSpaceBefore: 0,
            paraSpaceAfter: 0,
          },
          text: "(한국어 가사)",
        },
      },
      {
        placeholder: {
          options: {
            name: "engText", type: "body",
            x: LYRICS_ENG.x, y: LYRICS_ENG.y,
            w: LYRICS_ENG.w, h: LYRICS_ENG.h,
            fontSize: LYRICS_ENG.size,
            fontFace: LYRICS_ENG.font,
            color: LYRICS_ENG.color,
            align: LYRICS_ENG.align,
            valign: "top",
            lineSpacingMultiple: LYRICS_ENG.lineSpacing / 100,
            paraSpaceBefore: 0,
            paraSpaceAfter: 0,
          },
          text: "(English lyrics)",
        },
      },
    ],
  });

  // ── 마스터 2: 제목 슬라이드 ──
  pptx.defineSlideMaster({
    title: "TITLE_MASTER",
    background: { color: "000000" },
    objects: [
      bgObj,
      {
        placeholder: {
          options: {
            name: "songTitleKor", type: "title",
            x: TITLE_KOR.x, y: TITLE_KOR.y,
            w: TITLE_KOR.w, h: TITLE_KOR.h,
            fontSize: TITLE_KOR.size,
            fontFace: TITLE_KOR.font,
            color: TITLE_KOR.color,
            align: TITLE_KOR.align,
            valign: "bottom",
            lineSpacingMultiple: TITLE_KOR.lineSpacing / 100,
            paraSpaceBefore: 0,
            paraSpaceAfter: 0,
          },
          text: "한글 제목",
        },
      },
      // 구분선
      {
        line: {
          x: TITLE_LINE.x, y: TITLE_LINE.y,
          w: TITLE_LINE.w, h: 0,
          line: { color: "FFFFFF", width: TITLE_LINE.thickness },
        },
      },
      {
        placeholder: {
          options: {
            name: "songTitleEng", type: "body",
            x: TITLE_ENG.x, y: TITLE_ENG.y,
            w: TITLE_ENG.w, h: TITLE_ENG.h,
            fontSize: TITLE_ENG.size,
            fontFace: TITLE_ENG.font,
            color: TITLE_ENG.color,
            align: TITLE_ENG.align,
            valign: "top",
            lineSpacingMultiple: TITLE_ENG.lineSpacing / 100,
            paraSpaceBefore: 0,
            paraSpaceAfter: 0,
          },
          text: "English Title",
        },
      },
    ],
  });
}

// ─────────────────────────────────────────────────────────────────
//  가사 파싱
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
    } else if (line) {
      buf.push(line);
    }
  });
  if (buf.length) sections.push({ name, lines: [...buf] });
  return sections;
}

const isKor = t => /[ㄱ-ㅎㅏ-ㅣ가-힣]/.test(t);

// ─────────────────────────────────────────────────────────────────
//  PPT 생성
// ─────────────────────────────────────────────────────────────────
function generatePPT(title, rawLyrics) {
  if (!rawLyrics.trim()) { alert("가사를 입력해주세요."); return; }

  const pptx = new pptxgen();
  pptx.defineLayout({ name: "W20H11", width: 20, height: 11.25 });
  pptx.layout = "W20H11";
  defineMasters(pptx);

  const sections = parseLyrics(rawLyrics);

  // 1. 제목 슬라이드 (첫 번째 섹션)
  if (sections.length) {
    const sec = sections[0];
    pptx.addSection({ title: sec.name });
    const slide = pptx.addSlide({ masterName: "TITLE_MASTER", sectionTitle: sec.name });
    const korLine = sec.lines.find(isKor) || title || sec.name;
    const engLine = sec.lines.find(l => !isKor(l)) || "";
    slide.addText(korLine, { placeholder: "songTitleKor" });
    if (engLine) slide.addText(engLine, { placeholder: "songTitleEng" });
  }

  // 2. 가사 슬라이드 (나머지 섹션)
  sections.slice(1).forEach(sec => {
    pptx.addSection({ title: sec.name });
    const korLines = sec.lines.filter(isKor);
    const engLines = sec.lines.filter(l => !isKor(l));
    const max = Math.max(korLines.length, engLines.length, 1);

    for (let i = 0; i < max; i += 2) {
      const slide = pptx.addSlide({ masterName: "LYRICS_MASTER", sectionTitle: sec.name });
      const kor = korLines.slice(i, i + 2).join('\n');
      const eng = engLines.slice(i, i + 2).join('\n');
      if (kor) slide.addText(kor, { placeholder: "korText" });
      if (eng) slide.addText(eng, { placeholder: "engText" });
    }
  });

  pptx.writeFile({ fileName: `${title || "Worship_PPT"}.pptx` });
}

// ─────────────────────────────────────────────────────────────────
//  UI
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
    width: "100%", padding: "13px 16px", boxSizing: "border-box",
    background: "rgba(255,255,255,0.07)", border: "1px solid rgba(255,255,255,0.15)",
    borderRadius: "10px", color: "#0a0a0a", fontSize: "14px", outline: "none", fontFamily: "inherit",
  };

  return (
    <div style={{
      border: 0,
      minHeight: "100vh",
      display: "flex", alignItems: "center", justifyContent: "center",
      fontFamily: "'Segoe UI', sans-serif",
    }}>
      <div style={{
        width: "100%", maxWidth: "680px",
        backdropFilter: "blur(16px)",
        borderRadius: "20px", padding: "40px",
        border: "1px solid #0a0a0a",
      }}>
        {/* 헤더 */}
        <div style={{ textAlign: "center", marginBottom: "30px" }}>
          <p style={{ margin: 0, fontSize: "11px", letterSpacing: "4px", color: "#0a0a0a", textTransform: "uppercase" }}>
            Worship PPT Generator
          </p>
          <h1 style={{ margin: "8px 0 6px", fontSize: "26px", fontWeight: 700, color: "#0a0a0a" }}>
            찬양 PPT 만들기
          </h1>

        </div>

        {/* 파일 업로드 */}
        <label style={{
          display: "flex", alignItems: "center", gap: "12px", cursor: "pointer",
          background: "#f2f2f2", border: "1px dashed #0a0a0a",
          borderRadius: "10px", padding: "14px 18px", marginBottom: "14px",
        }}>
          <div>
            <div style={{ fontSize: "13px", color: "#0a0a0a", fontWeight: 600 }}>📤[TXT 파일 업로드]</div>
          </div>
          <input type="file" accept=".txt" onChange={handleFile} style={{ display: "none" }} />
        </label>

        {/* 제목 */}
        <input
          type="text"
          placeholder="찬양 제목"
          value={title}
          onChange={e => setTitle(e.target.value)}
          style={{ ...inputStyle, marginBottom: "12px", background: "#f2f2f2", border: "1px dashed #0a0a0a", }}
        />

        {/* 가사 */}
        <textarea
          placeholder={`가사를 입력하거나 파일을 올려주세요.\n\n예시:\n[이 산지를 내게 주소서]\n이 산지를 내게 주소서\nGive Me This Hill Country\n\n[V1]\n한국어 가사\nEnglish lyrics`}
          value={rawLyrics}
          onChange={e => setRawLyrics(e.target.value)}
          style={{ ...inputStyle, height: "300px", lineHeight: "1.65", resize: "vertical", marginBottom: "18px", background: "#f2f2f2", border: "1px dashed #0a0a0a", }}
        />

        {/* 버튼 */}
        <button
          onClick={() => generatePPT(title, rawLyrics)}
          style={{
            width: "100%", padding: "15px",
            color: "#0a0a0a", fontWeight: 700, fontSize: "15px",
            border: "1px solid rgba(109,191,94,0.35)", borderRadius: "10px",
            background: "#d1e8fb", border: "1px dashed #0a0a0a",
            cursor: "pointer", letterSpacing: "0.3px",
          }}
        >
          PPT 다운로드
        </button>
      </div>
    </div>
  );
}