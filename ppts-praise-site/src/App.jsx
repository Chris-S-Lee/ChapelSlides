import React, { useState, useRef } from 'react';
import pptxgen from "pptxgenjs";

// ─────────────────────────────────────────────────────────────────
// 1. 스타일 설정
// ─────────────────────────────────────────────────────────────────
const STYLE_CONFIG = {
  TITLE_KOR: {
    x: 2.83, y: 0.87, w: 17, h: 1.38,
    fontFace: "페이퍼로지 7 Bold", fontSize: 84, color: "FFFFFF",
    align: "right", valign: "bottom",
  },
  TITLE_ENG: {
    x: 2.83, y: 2.46, w: 17, h: 1.1,
    fontFace: "페이퍼로지 5 Medium", fontSize: 54, color: "FFFFFF",
    align: "right", valign: "top",
  },
  LYRICS_KOR: {
    x: 0, y: 0.38, w: 20, h: 2.2,
    fontFace: "페이퍼로지 5 Medium", fontSize: 72, color: "FFFFFF",
    align: "center", valign: "top", lineSpacingMultiple: 0.75,
  },
  LYRICS_ENG: {
    x: 0, y: 2.7, w: 20, h: 1.5,
    fontFace: "페이퍼로지 4 Regular", fontSize: 40, color: "FFFFFF",
    align: "center", valign: "top", lineSpacingMultiple: 0.7,
  },
};

// ─────────────────────────────────────────────────────────────────
// 2. 유틸리티
// ─────────────────────────────────────────────────────────────────
const isKor = t => /[ㄱ-ㅎㅏ-ㅣ가-힣]/.test(t);

function parseSections(raw) {
  const lines = raw.split('\n').map(l => l.trim());
  const sections = [];
  let name = null, buf = [];
  lines.forEach(line => {
    const tag = line.match(/^\[(.+)\]$/);
    if (tag) {
      if (name !== null) sections.push({ name, lines: [...buf] });
      else if (buf.length) sections.push({ name: "Intro", lines: [...buf] });
      name = tag[1]; buf = [];
    } else if (line) buf.push(line);
  });
  if (name !== null && buf.length) sections.push({ name, lines: [...buf] });
  else if (name === null && buf.length) sections.push({ name: "Intro", lines: buf });
  return sections;
}

function mergeSections(sections) {
  const merged = [];
  const seenNames = new Map();
  sections.forEach(sec => {
    const existing = seenNames.get(sec.name);
    if (existing !== undefined) {
      merged[existing].lines = [...merged[existing].lines, ...sec.lines];
    } else {
      seenNames.set(sec.name, merged.length);
      merged.push({ name: sec.name, lines: [...sec.lines] });
    }
  });
  return merged;
}

function defineMasters(pptx) {
  const bg_title = "/template_bg_title.png";
  const bg_content = "/template_bg_content.png";
  pptx.defineSlideMaster({
    title: "TITLE_MASTER",
    objects: [
      { image: { x: 0, y: 0, w: "100%", h: "100%", path: bg_title } },
      { placeholder: { options: { name: "songTitleKor", type: "title", ...STYLE_CONFIG.TITLE_KOR }, text: "" } },
      { placeholder: { options: { name: "songTitleEng", type: "body", ...STYLE_CONFIG.TITLE_ENG }, text: "" } },
    ],
  });
  pptx.defineSlideMaster({
    title: "LYRICS_MASTER",
    objects: [
      { image: { x: 0, y: 0, w: "100%", h: "100%", path: bg_content } },
      { placeholder: { options: { name: "korText", type: "body", ...STYLE_CONFIG.LYRICS_KOR }, text: "" } },
      { placeholder: { options: { name: "engText", type: "body", ...STYLE_CONFIG.LYRICS_ENG }, text: "" } },
    ],
  });
  pptx.defineSlideMaster({
    title: "BLANK_MASTER",
    objects: [
      { image: { x: 0, y: 0, w: "100%", h: "100%", path: bg_content } },
    ],
  });
}

function addSongToPPT(pptx, songTitle, sections, songIdx) {
  sections.forEach((sec, idx) => {
    const isLast = idx === sections.length - 1;
    const uniqueSecTitle = `${songIdx}__${sec.name}`;
    pptx.addSection({ title: uniqueSecTitle });

    if (idx === 0) {
      const slide = pptx.addSlide({ masterName: "TITLE_MASTER", sectionTitle: uniqueSecTitle });
      const korLines = sec.lines.filter(isKor);
      const engLines = sec.lines.filter(l => !isKor(l));
      const kor = korLines[0] || songTitle || sec.name;
      const eng = engLines[0] || "";
      slide.addText(kor, { placeholder: "songTitleKor", ...STYLE_CONFIG.TITLE_KOR });
      if (eng) slide.addText(eng, { placeholder: "songTitleEng", ...STYLE_CONFIG.TITLE_ENG });

      if (isLast) {
        pptx.addSlide({ masterName: "BLANK_MASTER", sectionTitle: uniqueSecTitle });
      }
    } else {
      const korLines = sec.lines.filter(isKor);
      const engLines = sec.lines.filter(l => !isKor(l));
      const maxLines = Math.max(korLines.length, engLines.length);

      for (let i = 0; i < maxLines; i += 2) {
        const slide = pptx.addSlide({ masterName: "LYRICS_MASTER", sectionTitle: uniqueSecTitle });
        const korPart = korLines.slice(i, i + 2).join('\n');
        const engPart = engLines.slice(i, i + 2).join('\n');
        if (korPart) slide.addText(korPart, { placeholder: "korText", ...STYLE_CONFIG.LYRICS_KOR });
        if (engPart) slide.addText(engPart, { placeholder: "engText", ...STYLE_CONFIG.LYRICS_ENG });
      }

      if (isLast) {
        pptx.addSlide({ masterName: "BLANK_MASTER", sectionTitle: uniqueSecTitle });
      }
    }
  });
}

async function generatePPT(songs) {
  if (!songs.length) return alert("파일을 추가해주세요.");
  const pptx = new pptxgen();
  pptx.defineLayout({ name: "W20H11", width: 20, height: 11.25 });
  pptx.layout = "W20H11";
  defineMasters(pptx);

  songs.forEach((song, songIdx) => {
    const raw      = parseSections(song.content);
    const merged   = mergeSections(raw);
    addSongToPPT(pptx, song.title, merged, songIdx);
  });

  const fileName = songs.length === 1 ? songs[0].title : "Worship_Combined";
  await pptx.writeFile({ fileName: `${fileName}.pptx` });
}

// ─────────────────────────────────────────────────────────────────
// 3. UI 설정
// ─────────────────────────────────────────────────────────────────
let uidCounter = 0;
const uid = () => `song_${++uidCounter}`;

const C = {
  bg:         "#f4f8f5",
  white:      "#ffffff",
  green:      "#4caf82",
  greenLight: "#e8f5ee",
  greenMid:   "#b2dfc5",
  border:     "#ddeae3",
  text:       "#1a2b22",
  sub:        "#5a7a68",
  muted:      "#9db8a8",
  danger:     "#e05c5c",
  shadow:     "rgba(76,175,130,0.10)",
};

export default function App() {
  const [songs, setSongs]       = useState([]);
  const [activeSongId, setActiveSongId] = useState(null);
  const [dragging, setDragging] = useState(null);
  const [dragOver, setDragOver] = useState(null);
  const fileInputRef = useRef();

  const handleFiles = e => {
    const files = Array.from(e.target.files);
    let firstNewId = null;

    files.forEach((f, index) => {
      const reader = new FileReader();
      const newId = uid();
      if (index === 0) firstNewId = newId;

      reader.onload = ev => {
        setSongs(prev => {
          const updated = [...prev, { id: newId, title: f.name.replace(/\.[^/.]+$/, ""), content: ev.target.result }];
          if (!activeSongId && index === 0) setActiveSongId(newId);
          return updated;
        });
      };
      reader.readAsText(f, "UTF-8");
    });
    e.target.value = "";
  };

  const updateTitle   = (id, val) => setSongs(prev => prev.map(s => s.id === id ? { ...s, title: val } : s));
  const updateContent = (id, val) => setSongs(prev => prev.map(s => s.id === id ? { ...s, content: val } : s));
  
  const removeSong  = id => {
    setSongs(prev => {
      const filtered = prev.filter(s => s.id !== id);
      if (activeSongId === id) {
        setActiveSongId(filtered.length ? filtered[0].id : null);
      }
      return filtered;
    });
  };

  const moveUp = id => setSongs(prev => {
    const i = prev.findIndex(s => s.id === id);
    if (i === 0) return prev;
    const a = [...prev]; [a[i-1], a[i]] = [a[i], a[i-1]]; return a;
  });
  const moveDown = id => setSongs(prev => {
    const i = prev.findIndex(s => s.id === id);
    if (i === prev.length - 1) return prev;
    const a = [...prev]; [a[i+1], a[i]] = [a[i], a[i+1]]; return a;
  });

  const onDragStart = (e, id) => { setDragging(id); e.dataTransfer.effectAllowed = "move"; };
  const onDragOver  = (e, id) => { e.preventDefault(); setDragOver(id); };
  const onDragEnd   = ()       => { setDragging(null); setDragOver(null); };
  const onDrop = (e, targetId) => {
    e.preventDefault();
    if (!dragging || dragging === targetId) { setDragging(null); setDragOver(null); return; }
    setSongs(prev => {
      const arr = [...prev];
      const fi = arr.findIndex(s => s.id === dragging);
      const ti = arr.findIndex(s => s.id === targetId);
      const [item] = arr.splice(fi, 1);
      arr.splice(ti, 0, item);
      return arr;
    });
    setDragging(null); setDragOver(null);
  };

  const activeSong = songs.find(s => s.id === activeSongId);

  return (
    <div style={{
      minHeight: "100vh", background: C.bg,
      fontFamily: "'Pretendard', 'Apple SD Gothic Neo', 'Noto Sans KR', sans-serif",
      padding: "52px 24px 80px",
    }}>
      {/* 항상 2단 그리드를 유지하도록 처리 */}
      <div style={{ 
        display: "grid", 
        gridTemplateColumns: "1fr 1fr", 
        gap: "32px", 
        maxWidth: "1280px", 
        margin: "0 auto",
        alignItems: "start"
      }}>
        
        {/* ================= 좌측 컴포넌트 (목록 및 설정) ================= */}
        <div>
          {/* 헤더 */}
          <div style={{ textAlign: "center", marginBottom: "36px" }}>
            <div style={{ fontSize: "26px", fontWeight: "800", color: C.text, letterSpacing: "-0.5px", marginBottom: "8px" }}>
              Worship PPT Generator
            </div>
          </div>

          {/* 업로드 박스 */}
          <div
            onClick={() => fileInputRef.current.click()}
            style={{
              background: C.white, border: `2px dashed ${C.greenMid}`,
              borderRadius: "18px", padding: "32px 20px",
              textAlign: "center", cursor: "pointer", marginBottom: "20px",
              transition: "background 0.15s, border-color 0.15s",
            }}
            onMouseEnter={e => { e.currentTarget.style.background = C.greenLight; e.currentTarget.style.borderColor = C.green; }}
            onMouseLeave={e => { e.currentTarget.style.background = C.white; e.currentTarget.style.borderColor = C.greenMid; }}
          >
            <div style={{ fontSize: "30px", marginBottom: "10px" }}>📂</div>
            <div style={{ color: C.green, fontWeight: "700", fontSize: "15px", marginBottom: "4px" }}>TXT 파일 추가</div>
            <div style={{ color: C.muted, fontSize: "12px" }}>여러 파일을 한 번에 선택할 수 있어요</div>
            <input ref={fileInputRef} type="file" accept=".txt" multiple onChange={handleFiles} style={{ display: "none" }} />
          </div>

          {/* 곡 목록 */}
          {songs.length > 0 ? (
            <>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "10px", padding: "0 2px" }}>
                <span style={{ fontSize: "12px", color: C.sub, fontWeight: "600" }}>
                  총 {songs.length}곡 · 드래그로 순서 변경 (클릭 시 가사 편집)
                </span>
                <span style={{ fontSize: "11px", color: C.muted }}>각 곡 끝에 빈 슬라이드 자동 삽입</span>
              </div>

              <div style={{ marginBottom: "20px" }}>
                {songs.map((song, i) => {
                  const isDragging = dragging === song.id;
                  const isOver     = dragOver === song.id && !isDragging;
                  const isActive   = activeSongId === song.id;

                  return (
                    <div
                      key={song.id}
                      draggable
                      onDragStart={e => onDragStart(e, song.id)}
                      onDragOver={e => onDragOver(e, song.id)}
                      onDrop={e => onDrop(e, song.id)}
                      onDragEnd={onDragEnd}
                      onClick={() => setActiveSongId(song.id)}
                      style={{
                        background: C.white,
                        border: `1.5px solid ${isOver ? C.green : isActive ? C.green : C.border}`,
                        borderRadius: "14px", padding: "12px 14px",
                        display: "flex", alignItems: "center", gap: "10px",
                        marginBottom: "8px", cursor: "pointer",
                        opacity: isDragging ? 0.4 : 1,
                        transform: isOver ? "scale(1.012)" : "scale(1)",
                        boxShadow: isActive 
                          ? `0 0 0 1px ${C.green}, 0 4px 12px ${C.shadow}` 
                          : isOver ? `0 0 0 3px ${C.greenLight}, 0 4px 16px ${C.shadow}` : `0 1px 4px ${C.shadow}`,
                        transition: "all 0.15s",
                      }}
                    >
                      <div style={{
                        width: "26px", height: "26px", borderRadius: "8px",
                        background: isActive ? C.green : C.greenLight, 
                        color: isActive ? C.white : C.green,
                        fontSize: "12px", fontWeight: "800",
                        display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0,
                      }}>{i + 1}</div>

                      <div style={{ color: C.muted, fontSize: "15px", flexShrink: 0, cursor: "grab", userSelect: "none" }} onClick={e => e.stopPropagation()}>⠿</div>

                      <input
                        type="text"
                        value={song.title}
                        onChange={e => updateTitle(song.id, e.target.value)}
                        onClick={e => e.stopPropagation()}
                        style={{
                          flex: 1, background: "none", border: "none", outline: "none",
                          color: C.text, fontSize: "14px", fontWeight: "600",
                          fontFamily: "Pretendard, inherit", minWidth: 0,
                        }}
                      />

                      {[
                        { label: "▲", fn: (e) => { e.stopPropagation(); moveUp(song.id); },   disabled: i === 0 },
                        { label: "▼", fn: (e) => { e.stopPropagation(); moveDown(song.id); }, disabled: i === songs.length - 1 },
                      ].map(btn => (
                        <button key={btn.label} onClick={btn.fn} disabled={btn.disabled}
                          style={{
                            background: "none", border: `1px solid ${C.border}`,
                            color: btn.disabled ? C.muted : C.sub,
                            borderRadius: "7px", width: "28px", height: "28px",
                            cursor: btn.disabled ? "default" : "pointer",
                            fontSize: "11px", display: "flex", alignItems: "center",
                            justifyContent: "center", flexShrink: 0,
                            opacity: btn.disabled ? 0.35 : 1,
                          }}
                        >{btn.label}</button>
                      ))}

                      <button onClick={(e) => { e.stopPropagation(); removeSong(song.id); }}
                        style={{
                          background: "none", border: `1px solid #f0c8c8`, color: C.danger,
                          borderRadius: "7px", width: "28px", height: "28px",
                          cursor: "pointer", fontSize: "13px", display: "flex",
                          alignItems: "center", justifyContent: "center", flexShrink: 0,
                        }}
                      >✕</button>
                    </div>
                  );
                })}
              </div>

              {/* 순서 미리보기 */}
              <div style={{
                background: C.white, border: `1px solid ${C.border}`,
                borderRadius: "14px", padding: "14px 18px", marginBottom: "20px",
              }}>
                <div style={{ fontSize: "12px", fontWeight: "700", color: C.sub, marginBottom: "10px" }}>
                  슬라이드 순서 미리보기
                </div>
                <div style={{ display: "flex", flexWrap: "wrap", gap: "6px", alignItems: "center" }}>
                  {songs.map((s, i) => (
                    <React.Fragment key={s.id}>
                      <span style={{
                        background: C.greenLight, color: C.green, border: `1px solid ${C.greenMid}`,
                        padding: "3px 11px", borderRadius: "20px", fontWeight: "700", fontSize: "12px",
                      }}>{s.title || `파일 ${i+1}`}</span>
                      {i < songs.length - 1 && <span style={{ color: C.muted }}>→</span>}
                    </React.Fragment>
                  ))}
                </div>
              </div>
            </>
          ) : (
            <div style={{ 
              textAlign: "center", padding: "40px 0", color: C.muted, fontSize: "14px",
              background: C.white, border: `1px solid ${C.border}`, borderRadius: "14px", marginBottom: "20px"
            }}>
              오른쪽 가이드를 참고하여 TXT 파일을 추가해 주세요.
            </div>
          )}

          {/* 생성 버튼 */}
          <button
            onClick={() => generatePPT(songs)}
            disabled={!songs.length}
            style={{
              width: "100%", padding: "16px",
              background: songs.length ? C.green : C.border,
              color: songs.length ? "#fff" : C.muted,
              border: "none", borderRadius: "14px",
              fontWeight: "800", fontSize: "16px",
              cursor: songs.length ? "pointer" : "not-allowed",
              letterSpacing: "-0.2px",
              boxShadow: songs.length ? `0 4px 20px ${C.shadow}` : "none",
              transition: "background 0.15s, transform 0.1s",
            }}
          >
            PPT 생성 및 다운로드
          </button>
        </div>

        {/* ================= 우측 컴포넌트 (처음부터 항상 노출) ================= */}
        <div style={{
          background: C.white, border: `1px solid ${C.border}`,
          borderRadius: "18px", padding: "24px", boxShadow: `0 4px 20px ${C.shadow}`,
          position: "sticky", top: "52px"
        }}>
          {activeSong ? (
            <>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "14px" }}>
                <div style={{ fontSize: "16px", fontWeight: "700", color: C.text }}>
                  📝 가사 미리보기 및 실시간 편집
                </div>
                <div style={{ fontSize: "12px", color: C.green, fontWeight: "600" }}>
                  {activeSong.title} 선택됨
                </div>
              </div>

              <textarea
                value={activeSong.content}
                onChange={e => updateContent(activeSong.id, e.target.value)}
                placeholder="여기에 가사 내용을 입력하거나 직접 수정할 수 있습니다."
                style={{
                  width: "100%", height: "32px", minHeight: "280px", padding: "16px",
                  border: `1.5px solid ${C.border}`, borderRadius: "12px",
                  fontSize: "13px", lineHeight: "1.6", color: C.text,
                  fontFamily: "Pretendard, 'Menlo', 'Consolas', monospace", resize: "vertical",
                  outline: "none", transition: "border-color 0.15s",
                  boxSizing: "border-box"
                }}
                onFocus={e => e.target.style.borderColor = C.green}
                onBlur={e => e.target.style.borderColor = C.border}
              />
            </>
          ) : (
            <div style={{ 
              color: C.text, fontSize: "15px", fontWeight: "700", marginBottom: "16px"
            }}>
              ℹ️ PPT 생성 안내 가이드
            </div>
          )}

          {/* 가이드 안내 영역 - 처음부터 항상 하단에 위치 */}
          <div style={{ marginTop: activeSong ? "24px" : "0px" }}>
            <div style={{ fontSize: "12px", fontWeight: "700", color: C.sub, marginBottom: "8px" }}>
              텍스트 파일 작성 규칙 예시:
            </div>
            <pre style={{
              color: C.sub, fontSize: "11px", lineHeight: "1.8", margin: 0,
              fontFamily: "Pretendard, 'Menlo', 'Consolas', monospace", whiteSpace: "pre-wrap",
              background: C.greenLight, padding: "12px 14px", borderRadius: "10px",
            }}>{`[Intro]\n찬양 제목 (한글)\nSong Title (English)\n\n[V]\n한글 가사 1행\n한글 가사 2행\n[C]\n한글 가사 3행\n한글 가사 4행`}</pre>
            <div style={{
              marginTop: "12px", padding: "10px 14px",
              background: C.greenLight, borderRadius: "10px",
              borderLeft: `3px solid ${C.green}`,
              fontSize: "11px", color: C.sub, lineHeight: "1.7",
            }}>
              💡 같은 섹션명([V], [C] 등)의 한글/영어 블록은 자동 병합됩니다.<br />
              💡 찬양 마지막에는 빈 슬라이드가 자동으로 포함됩니다.
            </div>
          </div>
        </div>

      </div>
    </div>
  );
}