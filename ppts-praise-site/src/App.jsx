import React, { useState } from 'react';
import pptxgen from "pptxgenjs";

function App() {
  const [title, setTitle] = useState('');
  const [rawLyrics, setRawLyrics] = useState('');

  // 사용자가 지정한 폰트명 (폴더에 있는 폰트 파일명과 일치해야 함)
  const CUSTOM_FONT = "YourCustomFontName"; // 실제 폰트 이름으로 수정하세요.

  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setTitle(file.name.replace(/\.[^/.]+$/, ""));
    const reader = new FileReader();
    reader.onload = (event) => setRawLyrics(event.target.result);
    reader.readAsText(file, "UTF-8");
  };

  const generatePPT = () => {
    let pptx = new pptxgen();
    pptx.layout = "LAYOUT_16x9";

    const lines = rawLyrics.split('\n').map(l => l.trim());
    let sections = [];
    let currentSectionName = "Intro";
    let currentBuffer = [];

    lines.forEach((line) => {
        const tagMatch = line.match(/^\[(.+)\]$/);
        if (tagMatch) {
            if (currentBuffer.length > 0) {
                sections.push({ title: currentSectionName, lines: [...currentBuffer] });
            }
            currentSectionName = tagMatch[1];
            currentBuffer = [];
        } else if (line === "" && currentBuffer.length > 0) {
            sections.push({ title: currentSectionName, lines: [...currentBuffer] });
            currentBuffer = [];
        } else if (line !== "") {
            currentBuffer.push(line);
        }
    });
    if (currentBuffer.length > 0) sections.push({ title: currentSectionName, lines: currentBuffer });

    const uniqueSectionNames = [...new Set(sections.map(s => s.title))];
    uniqueSectionNames.forEach(name => {
        pptx.addSection({ title: name });
    });

    let globalSlideIndex = 0;

    sections.forEach((section) => {
        const korLines = section.lines.filter(l => /[ㄱ-ㅎ|ㅏ-ㅣ|가-힣]/.test(l));
        const engLines = section.lines.filter(l => !/[ㄱ-ㅎ|ㅏ-ㅣ|가-힣]/.test(l));

        for (let i = 0; i < Math.max(korLines.length, engLines.length); i += 2) {
            let slide = pptx.addSlide({ sectionTitle: section.title });
            
            // 배경 설정: 기본 검정색 + 상단 1/3 어두운 초록색 사각형 (그라데이션 효과 모사)
            slide.background = { color: "000000" }; 
            slide.addShape(pptx.ShapeType.rect, {
                x: 0, y: 0, w: '100%', h: '33%',
                fill: { color: "062111" }
            });

            const korText = korLines.slice(i, i + 2).join('\n');
            const engText = engLines.slice(i, i + 2).join('\n');

            // 첫 페이지(제목 페이지) 여부 확인
            const isTitlePage = globalSlideIndex === 0;

            if (korText) {
                slide.addText(korText, {
                    x: 0.5, 
                    y: 0.8, 
                    w: '90%', 
                    h: 2.0,
                    fontSize: isTitlePage ? 72 : 32, // 제목 페이지는 2배
                    color: 'FFFFFF', 
                    align: isTitlePage ? 'right' : 'center', // 제목은 우측, 가사는 가운데
                    valign: 'top', // 모든 가사 상단 맞춤
                    bold: true, 
                    fontFace: CUSTOM_FONT
                });
            }
            if (engText) {
                slide.addText(engText, {
                    x: 0.5, 
                    y: korText ? 3.0 : 1.0, // 한국어 가사 유무에 따른 위치 조정
                    w: '90%', 
                    h: 1.5,
                    fontSize: isTitlePage ? 40 : 20, // 제목 페이지는 2배
                    color: 'FFFFFF', 
                    align: isTitlePage ? 'right' : 'center', 
                    valign: 'top',
                    italic: true, 
                    fontFace: CUSTOM_FONT
                });
            }
            globalSlideIndex++;
        }
    });

    pptx.writeFile({ fileName: `${title || '찬양PPT'}.pptx` });
  };

  return (
    <div style={{ padding: '40px', maxWidth: '700px', margin: '0 auto', fontFamily: 'sans-serif', backgroundColor: '#f9f9f9', borderRadius: '15px' }}>
      <h2 style={{ textAlign: 'center', color: '#062111' }}>Aesthetic Worship PPT Generator</h2>
      
      {/* ... (기존 UI 유지) */}
      <div style={{ border: '2px dashed #062111', padding: '20px', textAlign: 'center', marginBottom: '20px', borderRadius: '10px' }}>
        <input type="file" accept=".txt" onChange={handleFileUpload} />
      </div>

      <input 
        type="text" placeholder="찬양 제목" value={title}
        onChange={(e) => setTitle(e.target.value)}
        style={{ width: '100%', padding: '12px', marginBottom: '10px', borderRadius: '5px', border: '1px solid #ccc' }}
      />

      <textarea 
        placeholder="내용을 입력하세요." 
        value={rawLyrics}
        onChange={(e) => setRawLyrics(e.target.value)}
        style={{ width: '100%', height: '300px', padding: '15px', lineHeight: '1.6', borderRadius: '5px', border: '1px solid #ccc' }}
      />

      <button 
        onClick={generatePPT}
        style={{ width: '100%', padding: '15px', background: '#062111', color: '#fff', fontWeight: 'bold', cursor: 'pointer', border: 'none', borderRadius: '5px', marginTop: '10px' }}
      >
        새 디자인 PPT 다운로드
      </button>
    </div>
  );
}

export default App;