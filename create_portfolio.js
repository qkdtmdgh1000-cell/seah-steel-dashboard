const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");

// Icon imports
const {
  FaChartLine, FaChartBar, FaGlobeAsia, FaLightbulb,
  FaCalendarAlt, FaBullseye, FaArrowUp, FaArrowDown,
  FaIndustry, FaFileAlt, FaChartPie, FaUsers,
  FaMoneyBillWave, FaCogs, FaCheckCircle, FaExclamationTriangle,
  FaBalanceScale, FaRocket, FaClipboardList, FaSearchDollar
} = require("react-icons/fa");

// === ICON UTILITY ===
function renderIconSvg(IconComponent, color = "#000000", size = 256) {
  return ReactDOMServer.renderToStaticMarkup(
    React.createElement(IconComponent, { color, size: String(size) })
  );
}

async function iconToBase64Png(IconComponent, color, size = 256) {
  const svg = renderIconSvg(IconComponent, color, size);
  const pngBuffer = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + pngBuffer.toString("base64");
}

// === COLOR PALETTE (Steel Industry — Deep Navy + Steel Blue) ===
const C = {
  navy:       "1B2A4A",
  darkNavy:   "0F1B32",
  steelBlue:  "3B6B9C",
  accent:     "2E86AB",
  lightBlue:  "A8D0E6",
  paleBlue:   "E8F1F8",
  white:      "FFFFFF",
  offWhite:   "F5F7FA",
  lightGray:  "E2E8F0",
  midGray:    "94A3B8",
  darkGray:   "475569",
  textDark:   "1E293B",
  textMid:    "64748B",
  red:        "DC2626",
  green:      "16A34A",
  orange:     "EA580C",
  amber:      "D97706",
  teal:       "0D9488",
};

// === HELPER FUNCTIONS ===
const makeShadow = () => ({
  type: "outer", blur: 6, offset: 2, angle: 135, color: "000000", opacity: 0.12
});

const makeCardShadow = () => ({
  type: "outer", blur: 8, offset: 3, angle: 135, color: "000000", opacity: 0.10
});

function addFooter(slide, pageNum, totalPages) {
  slide.addText(`${pageNum} / ${totalPages}`, {
    x: 8.5, y: 5.2, w: 1.2, h: 0.3,
    fontSize: 9, color: C.midGray, align: "right", fontFace: "Calibri"
  });
  slide.addText("경영기획 직무역량 포트폴리오", {
    x: 0.5, y: 5.2, w: 4, h: 0.3,
    fontSize: 9, color: C.midGray, align: "left", fontFace: "Calibri"
  });
}

function addSectionHeader(slide, sectionNum, sectionTitle) {
  slide.addShape(slide._slideLayout ? "rect" : "rect", {}); // dummy
  // Actually use proper shape
}

async function main() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "지원자";
  pres.title = "세아제강 경영기획 직무역량 포트폴리오";

  const TOTAL_PAGES = 12;

  // Pre-render icons
  const icons = {
    chartLine:   await iconToBase64Png(FaChartLine, `#${C.white}`),
    chartBar:    await iconToBase64Png(FaChartBar, `#${C.white}`),
    globe:       await iconToBase64Png(FaGlobeAsia, `#${C.white}`),
    lightbulb:   await iconToBase64Png(FaLightbulb, `#${C.white}`),
    calendar:    await iconToBase64Png(FaCalendarAlt, `#${C.white}`),
    bullseye:    await iconToBase64Png(FaBullseye, `#${C.white}`),
    arrowUp:     await iconToBase64Png(FaArrowUp, `#${C.green}`),
    arrowDown:   await iconToBase64Png(FaArrowDown, `#${C.red}`),
    industry:    await iconToBase64Png(FaIndustry, `#${C.white}`),
    fileAlt:     await iconToBase64Png(FaFileAlt, `#${C.white}`),
    chartPie:    await iconToBase64Png(FaChartPie, `#${C.white}`),
    users:       await iconToBase64Png(FaUsers, `#${C.white}`),
    money:       await iconToBase64Png(FaMoneyBillWave, `#${C.white}`),
    cogs:        await iconToBase64Png(FaCogs, `#${C.white}`),
    check:       await iconToBase64Png(FaCheckCircle, `#${C.green}`),
    warning:     await iconToBase64Png(FaExclamationTriangle, `#${C.orange}`),
    balance:     await iconToBase64Png(FaBalanceScale, `#${C.white}`),
    rocket:      await iconToBase64Png(FaRocket, `#${C.white}`),
    clipboard:   await iconToBase64Png(FaClipboardList, `#${C.white}`),
    searchDollar:await iconToBase64Png(FaSearchDollar, `#${C.white}`),
    // Dark versions for light backgrounds
    chartLineDk: await iconToBase64Png(FaChartLine, `#${C.steelBlue}`),
    globeDk:     await iconToBase64Png(FaGlobeAsia, `#${C.steelBlue}`),
    lightbulbDk: await iconToBase64Png(FaLightbulb, `#${C.steelBlue}`),
    calendarDk:  await iconToBase64Png(FaCalendarAlt, `#${C.steelBlue}`),
    industryDk:  await iconToBase64Png(FaIndustry, `#${C.steelBlue}`),
    moneyDk:     await iconToBase64Png(FaMoneyBillWave, `#${C.steelBlue}`),
    cogsDk:      await iconToBase64Png(FaCogs, `#${C.steelBlue}`),
    clipboardDk: await iconToBase64Png(FaClipboardList, `#${C.steelBlue}`),
    rocketDk:    await iconToBase64Png(FaRocket, `#${C.steelBlue}`),
    checkGreen:  await iconToBase64Png(FaCheckCircle, `#${C.green}`),
    balanceDk:   await iconToBase64Png(FaBalanceScale, `#${C.steelBlue}`),
    searchDk:    await iconToBase64Png(FaSearchDollar, `#${C.steelBlue}`),
    pieDk:       await iconToBase64Png(FaChartPie, `#${C.steelBlue}`),
    barDk:       await iconToBase64Png(FaChartBar, `#${C.steelBlue}`),
  };

  // ============================================================
  // SLIDE 1: COVER
  // ============================================================
  {
    const slide = pres.addSlide();
    slide.background = { color: C.darkNavy };

    // Decorative top bar
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent }
    });

    // Left accent stripe
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y: 1.2, w: 0.08, h: 2.8, fill: { color: C.accent }
    });

    // Title
    slide.addText("경영기획 직무역량\n포트폴리오", {
      x: 0.9, y: 1.2, w: 6, h: 2.0,
      fontSize: 38, fontFace: "Calibri", color: C.white, bold: true,
      lineSpacingMultiple: 1.3, margin: 0
    });

    // Subtitle
    slide.addText("데이터 기반 분석 및 보고서 작성 역량 중심", {
      x: 0.9, y: 3.1, w: 6, h: 0.5,
      fontSize: 16, fontFace: "Calibri", color: C.lightBlue, margin: 0
    });

    // Info box
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.9, y: 3.9, w: 3.5, h: 0.9,
      fill: { color: C.navy }, rectRadius: 0
    });
    slide.addText([
      { text: "지원 직무  ", options: { color: C.midGray, fontSize: 11, breakLine: true } },
      { text: "세아제강 경영기획 신입사원", options: { color: C.white, fontSize: 13, bold: true } }
    ], { x: 1.1, y: 3.95, w: 3.2, h: 0.8, fontFace: "Calibri", margin: 0 });

    // Right side decorative icon area
    slide.addShape(pres.shapes.OVAL, {
      x: 7.2, y: 1.5, w: 2.2, h: 2.2,
      fill: { color: C.steelBlue, transparency: 30 }
    });
    slide.addShape(pres.shapes.OVAL, {
      x: 7.6, y: 1.9, w: 1.4, h: 1.4,
      fill: { color: C.accent, transparency: 20 }
    });
    slide.addImage({ data: icons.chartLine, x: 7.95, y: 2.2, w: 0.7, h: 0.7 });

    // Bottom date
    slide.addText("2026년 3월", {
      x: 7, y: 4.8, w: 2.5, h: 0.4,
      fontSize: 12, fontFace: "Calibri", color: C.midGray, align: "right"
    });
  }

  // ============================================================
  // SLIDE 2: TABLE OF CONTENTS
  // ============================================================
  {
    const slide = pres.addSlide();
    slide.background = { color: C.offWhite };

    // Top colored bar
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent }
    });

    slide.addText("목차", {
      x: 0.6, y: 0.35, w: 4, h: 0.6,
      fontSize: 28, fontFace: "Calibri", color: C.textDark, bold: true, margin: 0
    });
    slide.addText("CONTENTS", {
      x: 0.6, y: 0.85, w: 4, h: 0.35,
      fontSize: 11, fontFace: "Calibri", color: C.midGray, charSpacing: 4, margin: 0
    });

    const tocItems = [
      { num: "01", title: "연도별 경영실적 분석 보고", desc: "KPI 대시보드 · 매출 추이 · 수익성 분석", icon: icons.barDk },
      { num: "02", title: "사업환경 변화 분석 및 대응 방안", desc: "원자재 시황 · 글로벌 트렌드 · SWOT 분석", icon: icons.globeDk },
      { num: "03", title: "중장기 경영계획 수립", desc: "3개년 전략 로드맵", icon: icons.rocketDk },
      { num: "04", title: "투자예산 및 회의체 운영관리", desc: "예산 집행 현황 대시보드", icon: icons.clipboardDk },
    ];

    tocItems.forEach((item, i) => {
      const yBase = 1.45 + i * 0.78;

      // Card background
      slide.addShape(pres.shapes.RECTANGLE, {
        x: 0.6, y: yBase, w: 8.8, h: 0.68,
        fill: { color: C.white }, shadow: makeCardShadow()
      });

      // Left accent
      slide.addShape(pres.shapes.RECTANGLE, {
        x: 0.6, y: yBase, w: 0.06, h: 0.68,
        fill: { color: i < 2 ? C.accent : C.steelBlue }
      });

      // Number
      slide.addText(item.num, {
        x: 0.85, y: yBase, w: 0.55, h: 0.68,
        fontSize: 20, fontFace: "Calibri", color: i < 2 ? C.accent : C.steelBlue,
        bold: true, valign: "middle", margin: 0
      });

      // Icon
      slide.addImage({ data: item.icon, x: 1.45, y: yBase + 0.17, w: 0.34, h: 0.34 });

      // Title
      slide.addText(item.title, {
        x: 2.0, y: yBase + 0.05, w: 5, h: 0.35,
        fontSize: 14, fontFace: "Calibri", color: C.textDark, bold: true, margin: 0
      });

      // Description
      slide.addText(item.desc, {
        x: 2.0, y: yBase + 0.35, w: 5, h: 0.3,
        fontSize: 10, fontFace: "Calibri", color: C.textMid, margin: 0
      });

      // Priority badge for items 0,1
      if (i < 2) {
        slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
          x: 8.0, y: yBase + 0.18, w: 1.15, h: 0.32,
          fill: { color: C.accent }, rectRadius: 0.05
        });
        slide.addText("중점 역량", {
          x: 8.0, y: yBase + 0.18, w: 1.15, h: 0.32,
          fontSize: 9, fontFace: "Calibri", color: C.white, bold: true,
          align: "center", valign: "middle"
        });
      }
    });

    addFooter(slide, 2, TOTAL_PAGES);
  }

  // ============================================================
  // SLIDE 3: SECTION DIVIDER — 월별 경영실적 분석
  // ============================================================
  {
    const slide = pres.addSlide();
    slide.background = { color: C.navy };

    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent }
    });

    slide.addText("01", {
      x: 0.8, y: 1.2, w: 2, h: 1.0,
      fontSize: 60, fontFace: "Calibri", color: C.accent, bold: true, margin: 0
    });
    slide.addText("연도별 경영실적 분석 보고", {
      x: 0.8, y: 2.2, w: 8, h: 0.7,
      fontSize: 30, fontFace: "Calibri", color: C.white, bold: true, margin: 0
    });
    slide.addText("Annual Business Performance Analysis", {
      x: 0.8, y: 2.9, w: 8, h: 0.4,
      fontSize: 13, fontFace: "Calibri", color: C.lightBlue, margin: 0
    });

    // Icon circle
    slide.addShape(pres.shapes.OVAL, {
      x: 7.8, y: 1.8, w: 1.5, h: 1.5,
      fill: { color: C.steelBlue, transparency: 40 }
    });
    slide.addImage({ data: icons.chartBar, x: 8.15, y: 2.15, w: 0.8, h: 0.8 });

    // Bottom description
    slide.addText("세아제강 사업보고서(2023~2025) 실제 데이터를 기반으로\nKPI 분석, 매출 추이, 수익성 분석 보고서를 작성합니다.", {
      x: 0.8, y: 3.7, w: 7, h: 0.8,
      fontSize: 12, fontFace: "Calibri", color: C.midGray, lineSpacingMultiple: 1.5, margin: 0
    });

    addFooter(slide, 3, TOTAL_PAGES);
  }

  // ============================================================
  // SLIDE 4: KPI DASHBOARD
  // ============================================================
  {
    const slide = pres.addSlide();
    slide.background = { color: C.offWhite };
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent }
    });

    slide.addText("핵심 경영지표 대시보드 (2025년 실적)", {
      x: 0.5, y: 0.25, w: 7, h: 0.5,
      fontSize: 20, fontFace: "Calibri", color: C.textDark, bold: true, margin: 0
    });
    slide.addText("KPI Dashboard  |  출처: 사업보고서(제8기)", {
      x: 0.5, y: 0.7, w: 5, h: 0.3,
      fontSize: 10, fontFace: "Calibri", color: C.midGray, margin: 0
    });

    // KPI Cards — 4 across (실제 데이터: 사업보고서 제8기/2025)
    const kpis = [
      { label: "매출액", value: "14,848", unit: "억원", change: "-17.9%", up: false },
      { label: "영업이익률", value: "3.3", unit: "%", change: "-7.9%p", up: false },
      { label: "생산실적", value: "79.4", unit: "만톤", change: "능력 160만톤", up: false },
      { label: "평균 가동률", value: "84", unit: "%", change: "창원 93% 최고", up: true },
    ];

    kpis.forEach((kpi, i) => {
      const x = 0.5 + i * 2.3;
      const y = 1.2;
      const w = 2.1;

      // Card
      slide.addShape(pres.shapes.RECTANGLE, {
        x, y, w, h: 1.45, fill: { color: C.white }, shadow: makeCardShadow()
      });
      // Top accent
      slide.addShape(pres.shapes.RECTANGLE, {
        x, y, w, h: 0.05, fill: { color: kpi.up ? C.accent : C.orange }
      });

      slide.addText(kpi.label, {
        x: x + 0.15, y: y + 0.12, w: w - 0.3, h: 0.3,
        fontSize: 10, fontFace: "Calibri", color: C.textMid, margin: 0
      });
      slide.addText([
        { text: kpi.value, options: { fontSize: 28, bold: true, color: C.textDark } },
        { text: ` ${kpi.unit}`, options: { fontSize: 12, color: C.textMid } }
      ], {
        x: x + 0.15, y: y + 0.4, w: w - 0.3, h: 0.5,
        fontFace: "Calibri", margin: 0
      });

      // Change badge
      slide.addImage({
        data: kpi.up ? icons.arrowUp : icons.arrowDown,
        x: x + 0.15, y: y + 1.0, w: 0.18, h: 0.18
      });
      slide.addText(`전월 대비 ${kpi.change}`, {
        x: x + 0.38, y: y + 0.95, w: w - 0.6, h: 0.3,
        fontSize: 10, fontFace: "Calibri", color: kpi.up ? C.green : C.red, margin: 0
      });
    });

    // CHART: 3개년 연결 실적 추이 (실제 사업보고서 데이터)
    slide.addText("3개년 연결 실적 추이 (제6~8기)", {
      x: 0.5, y: 2.9, w: 5, h: 0.4,
      fontSize: 13, fontFace: "Calibri", color: C.textDark, bold: true, margin: 0
    });

    slide.addChart(pres.charts.BAR, [
      {
        name: "매출액(억원)",
        labels: ["2023(제6기)", "2024(제7기)", "2025(제8기)"],
        values: [18609, 18094, 14848]
      },
      {
        name: "영업이익(억원)",
        labels: ["2023(제6기)", "2024(제7기)", "2025(제8기)"],
        values: [2319, 2029, 496]
      }
    ], {
      x: 0.3, y: 3.2, w: 5.2, h: 2.1, barDir: "col",
      chartColors: [C.steelBlue, C.accent],
      chartArea: { fill: { color: C.white }, roundedCorners: true },
      catAxisLabelColor: C.textMid, catAxisLabelFontSize: 8,
      valAxisLabelColor: C.textMid, valAxisLabelFontSize: 8,
      valGridLine: { color: C.lightGray, size: 0.5 },
      catGridLine: { style: "none" },
      showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 8,
      showLegend: true, legendPos: "b", legendFontSize: 9,
    });

    // Right side: Key Insights card
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 5.8, y: 2.9, w: 3.9, h: 2.35,
      fill: { color: C.white }, shadow: makeCardShadow()
    });
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 5.8, y: 2.9, w: 0.06, h: 2.35,
      fill: { color: C.accent }
    });
    slide.addText("분석 인사이트", {
      x: 6.1, y: 3.0, w: 3.4, h: 0.35,
      fontSize: 12, fontFace: "Calibri", color: C.accent, bold: true, margin: 0
    });
    slide.addText([
      { text: "1. 매출 14,848억 → 전년比 17.9% 감소", options: { bullet: true, breakLine: true, fontSize: 10, color: C.textDark } },
      { text: "2. 영업이익률 3.3% → 수익성 급락 (전년 11.2%)", options: { bullet: true, breakLine: true, fontSize: 10, color: C.textDark } },
      { text: "3. 매출원가율 90.3% → 원가 부담 심화", options: { bullet: true, breakLine: true, fontSize: 10, color: C.textDark } },
      { text: "4. HR Coil 가격 하락 (885→809천원/톤)", options: { bullet: true, breakLine: true, fontSize: 10, color: C.textDark } },
      { text: "5. 창원 가동률 93%로 최고 효율 유지", options: { bullet: true, fontSize: 10, color: C.textDark } },
    ], {
      x: 6.1, y: 3.4, w: 3.4, h: 1.7,
      fontFace: "Calibri", margin: 0, paraSpaceAfter: 6
    });

    addFooter(slide, 4, TOTAL_PAGES);
  }

  // ============================================================
  // SLIDE 5: PROFITABILITY ANALYSIS
  // ============================================================
  {
    const slide = pres.addSlide();
    slide.background = { color: C.offWhite };
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent }
    });

    slide.addText("수익성 분석 및 원가 구조", {
      x: 0.5, y: 0.25, w: 7, h: 0.5,
      fontSize: 20, fontFace: "Calibri", color: C.textDark, bold: true, margin: 0
    });

    // Line chart: 3개년 이익률 추이 (실제 데이터)
    slide.addChart(pres.charts.LINE, [
      {
        name: "매출총이익률(%)",
        labels: ["2023(제6기)", "2024(제7기)", "2025(제8기)"],
        values: [17.1, 16.5, 9.7]
      },
      {
        name: "영업이익률(%)",
        labels: ["2023(제6기)", "2024(제7기)", "2025(제8기)"],
        values: [12.5, 11.2, 3.3]
      },
      {
        name: "순이익률(%)",
        labels: ["2023(제6기)", "2024(제7기)", "2025(제8기)"],
        values: [10.1, 7.6, 2.0]
      }
    ], {
      x: 0.3, y: 0.9, w: 5.5, h: 2.5,
      chartColors: [C.accent, C.steelBlue, C.orange],
      lineSmooth: true, lineSize: 2.5,
      chartArea: { fill: { color: C.white }, roundedCorners: true },
      catAxisLabelColor: C.textMid, catAxisLabelFontSize: 8,
      valAxisLabelColor: C.textMid, valAxisLabelFontSize: 8,
      valGridLine: { color: C.lightGray, size: 0.5 },
      catGridLine: { style: "none" },
      showLegend: true, legendPos: "b", legendFontSize: 9,
    });

    // 매출 구성비 (2025 실제 데이터)
    slide.addText("매출 원가 구조 (2025)", {
      x: 6.2, y: 0.8, w: 3, h: 0.35,
      fontSize: 13, fontFace: "Calibri", color: C.textDark, bold: true, margin: 0, align: "center"
    });

    slide.addChart(pres.charts.DOUGHNUT, [{
      name: "매출구조",
      labels: ["매출원가", "판관비", "영업이익"],
      values: [90.3, 6.3, 3.3]
    }], {
      x: 6.3, y: 1.1, w: 3.2, h: 2.3,
      chartColors: [C.navy, C.steelBlue, C.accent],
      showPercent: true, showLegend: true, legendPos: "b", legendFontSize: 8,
      dataLabelColor: C.white, dataLabelFontSize: 9,
    });

    // 3개년 손익 비교 테이블 (실제 사업보고서 데이터, 단위: 억원)
    slide.addText("3개년 손익 비교 (연결 기준, 단위: 억원)", {
      x: 0.5, y: 3.55, w: 5, h: 0.35,
      fontSize: 13, fontFace: "Calibri", color: C.textDark, bold: true, margin: 0
    });

    const tableHeader = [
      [
        { text: "구분", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 10, align: "center" } },
        { text: "매출액", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 10, align: "center" } },
        { text: "매출원가", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 10, align: "center" } },
        { text: "판관비", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 10, align: "center" } },
        { text: "영업이익", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 10, align: "center" } },
        { text: "당기순이익", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 10, align: "center" } },
      ],
      [
        { text: "2023(제6기)", options: { fill: { color: C.white }, fontSize: 10, align: "center" } },
        { text: "18,609", options: { fill: { color: C.white }, fontSize: 10, align: "center" } },
        { text: "15,435", options: { fill: { color: C.white }, fontSize: 10, align: "center" } },
        { text: "855", options: { fill: { color: C.white }, fontSize: 10, align: "center" } },
        { text: "2,319", options: { fill: { color: C.white }, fontSize: 10, align: "center", bold: true } },
        { text: "1,888", options: { fill: { color: C.white }, fontSize: 10, align: "center" } },
      ],
      [
        { text: "2024(제7기)", options: { fill: { color: C.paleBlue }, fontSize: 10, align: "center" } },
        { text: "18,094", options: { fill: { color: C.paleBlue }, fontSize: 10, align: "center" } },
        { text: "15,110", options: { fill: { color: C.paleBlue }, fontSize: 10, align: "center" } },
        { text: "955", options: { fill: { color: C.paleBlue }, fontSize: 10, align: "center" } },
        { text: "2,029", options: { fill: { color: C.paleBlue }, fontSize: 10, align: "center", bold: true } },
        { text: "1,371", options: { fill: { color: C.paleBlue }, fontSize: 10, align: "center" } },
      ],
      [
        { text: "2025(제8기)", options: { fill: { color: C.white }, fontSize: 10, align: "center", bold: true } },
        { text: "14,848", options: { fill: { color: C.white }, fontSize: 10, align: "center", color: C.red } },
        { text: "13,415", options: { fill: { color: C.white }, fontSize: 10, align: "center" } },
        { text: "937", options: { fill: { color: C.white }, fontSize: 10, align: "center" } },
        { text: "496", options: { fill: { color: C.white }, fontSize: 10, align: "center", color: C.red, bold: true } },
        { text: "300", options: { fill: { color: C.white }, fontSize: 10, align: "center", color: C.red } },
      ],
    ];

    slide.addTable(tableHeader, {
      x: 0.5, y: 3.9, w: 9, h: 1.3,
      border: { pt: 0.5, color: C.lightGray },
      colW: [1.2, 1.6, 1.5, 1.5, 1.5, 1.7],
      fontFace: "Calibri",
    });

    addFooter(slide, 5, TOTAL_PAGES);
  }

  // ============================================================
  // SLIDE 6: EXECUTIVE SUMMARY (Monthly Analysis)
  // ============================================================
  {
    const slide = pres.addSlide();
    slide.background = { color: C.offWhite };
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent }
    });

    slide.addText("2025년 실적 종합 평가 및 향후 전망", {
      x: 0.5, y: 0.25, w: 7, h: 0.5,
      fontSize: 20, fontFace: "Calibri", color: C.textDark, bold: true, margin: 0
    });

    // Left: Summary card
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y: 1.0, w: 4.5, h: 4.0,
      fill: { color: C.white }, shadow: makeCardShadow()
    });
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y: 1.0, w: 4.5, h: 0.5,
      fill: { color: C.navy }
    });
    slide.addText("2025년(제8기) 종합 평가", {
      x: 0.7, y: 1.05, w: 4, h: 0.4,
      fontSize: 14, fontFace: "Calibri", color: C.white, bold: true, margin: 0
    });

    const summaryItems = [
      { icon: icons.warning, text: "매출 14,848억원 → 전년比 17.9% 감소", color: C.orange },
      { icon: icons.warning, text: "영업이익률 3.3% → 전년比 7.9%p 급락", color: C.orange },
      { icon: icons.warning, text: "매출원가율 90.3% → 원가 부담 가중", color: C.orange },
      { icon: icons.checkGreen, text: "EBITDA 848억원 → 현금창출력 유지", color: C.green },
      { icon: icons.checkGreen, text: "배당성향 52%, 주당 5,500원 주주환원 지속", color: C.green },
    ];

    summaryItems.forEach((item, i) => {
      const y = 1.7 + i * 0.55;
      slide.addImage({ data: item.icon, x: 0.8, y: y + 0.02, w: 0.25, h: 0.25 });
      slide.addText(item.text, {
        x: 1.2, y: y, w: 3.6, h: 0.35,
        fontSize: 11, fontFace: "Calibri", color: C.textDark, margin: 0, valign: "middle"
      });
    });

    // Overall rating
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.8, y: 4.3, w: 3.9, h: 0.5,
      fill: { color: C.paleBlue }
    });
    slide.addText("종합 등급:  C  (개선 필요) — 수익성 회복 시급", {
      x: 0.8, y: 4.3, w: 3.9, h: 0.5,
      fontSize: 12, fontFace: "Calibri", color: C.red, bold: true,
      align: "center", valign: "middle"
    });

    // Right: Next month outlook
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 5.3, y: 1.0, w: 4.3, h: 4.0,
      fill: { color: C.white }, shadow: makeCardShadow()
    });
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 5.3, y: 1.0, w: 4.3, h: 0.5,
      fill: { color: C.steelBlue }
    });
    slide.addText("2026년 전망 및 핵심 과제", {
      x: 5.5, y: 1.05, w: 4, h: 0.4,
      fontSize: 14, fontFace: "Calibri", color: C.white, bold: true, margin: 0
    });

    const outlookItems = [
      { title: "수익성 회복", text: "원가율 90.3% → 85% 이하 목표\n고부가 제품 믹스 확대, 원재료 조달 다변화" },
      { title: "리스크 요인", text: "철광석 $100/톤, 철스크랩 $330~370 변동성\n원/달러 환율 1,421원대 고환율 지속" },
      { title: "성장 동력", text: "SeAH Wind 해상풍력 투자 (RCPS 1,479억원)\n해외법인 확대 (SSUSA·SSV 등 7개국 81만톤)" },
    ];

    outlookItems.forEach((item, i) => {
      const y = 1.7 + i * 0.78;
      slide.addText(item.title, {
        x: 5.6, y, w: 3.8, h: 0.25,
        fontSize: 11, fontFace: "Calibri", color: C.accent, bold: true, margin: 0
      });
      slide.addText(item.text, {
        x: 5.6, y: y + 0.25, w: 3.8, h: 0.5,
        fontSize: 10, fontFace: "Calibri", color: C.textDark, margin: 0, lineSpacingMultiple: 1.3
      });
    });

    addFooter(slide, 6, TOTAL_PAGES);
  }

  // ============================================================
  // SLIDE 7: SECTION DIVIDER — 사업환경 변화 분석
  // ============================================================
  {
    const slide = pres.addSlide();
    slide.background = { color: C.navy };
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent }
    });

    slide.addText("02", {
      x: 0.8, y: 1.2, w: 2, h: 1.0,
      fontSize: 60, fontFace: "Calibri", color: C.accent, bold: true, margin: 0
    });
    slide.addText("사업환경 변화 분석 및\n대응 방안 수립", {
      x: 0.8, y: 2.2, w: 8, h: 0.9,
      fontSize: 28, fontFace: "Calibri", color: C.white, bold: true, margin: 0,
      lineSpacingMultiple: 1.2
    });
    slide.addText("Business Environment Analysis & Response Strategy", {
      x: 0.8, y: 3.1, w: 8, h: 0.4,
      fontSize: 13, fontFace: "Calibri", color: C.lightBlue, margin: 0
    });

    slide.addShape(pres.shapes.OVAL, {
      x: 7.8, y: 1.8, w: 1.5, h: 1.5,
      fill: { color: C.steelBlue, transparency: 40 }
    });
    slide.addImage({ data: icons.globe, x: 8.15, y: 2.15, w: 0.8, h: 0.8 });

    slide.addText("철광석·철스크랩 가격, 건설투자, 중국산 수입 동향 등 실제 시장 데이터를 분석하고\n세아제강의 전략적 대응 방안을 수립합니다.", {
      x: 0.8, y: 3.7, w: 7, h: 0.8,
      fontSize: 12, fontFace: "Calibri", color: C.midGray, lineSpacingMultiple: 1.5, margin: 0
    });

    addFooter(slide, 7, TOTAL_PAGES);
  }

  // ============================================================
  // SLIDE 8: MARKET TRENDS
  // ============================================================
  {
    const slide = pres.addSlide();
    slide.background = { color: C.offWhite };
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent }
    });

    slide.addText("글로벌 철강 시장 동향", {
      x: 0.5, y: 0.25, w: 7, h: 0.5,
      fontSize: 20, fontFace: "Calibri", color: C.textDark, bold: true, margin: 0
    });

    // Chart: Raw material price trends
    slide.addText("주요 원자재 가격 추이 (USD/톤)", {
      x: 0.5, y: 0.85, w: 5, h: 0.3,
      fontSize: 12, fontFace: "Calibri", color: C.textMid, margin: 0
    });

    slide.addChart(pres.charts.LINE, [
      {
        name: "철광석($/dmt)",
        labels: ["2024 Q1", "2024 Q2", "2024 Q3", "2024 Q4", "2025 평균"],
        values: [130, 119, 100, 100, 100]
      },
      {
        name: "철스크랩($/톤)",
        labels: ["2024 Q1", "2024 Q2", "2024 Q3", "2024 Q4", "2025 평균"],
        values: [371, 350, 341, 322, 350]
      },
    ], {
      x: 0.3, y: 1.1, w: 5.2, h: 2.2,
      chartColors: [C.accent, C.orange],
      lineSmooth: true, lineSize: 2.5,
      chartArea: { fill: { color: C.white }, roundedCorners: true },
      catAxisLabelColor: C.textMid, catAxisLabelFontSize: 8,
      valAxisLabelColor: C.textMid, valAxisLabelFontSize: 8,
      valGridLine: { color: C.lightGray, size: 0.5 },
      catGridLine: { style: "none" },
      showLegend: true, legendPos: "b", legendFontSize: 9,
    });

    // Right side: Key trends cards
    const trends = [
      { title: "국내 건설투자", text: "2024 290.2조(0.0%)\n2025 264.7조(-8.8%) 감소", color: C.red },
      { title: "중국산 수입 변화", text: "2024 877만톤 → 2025 620만톤\n반덤핑 관세 효과 가시화", color: C.green },
      { title: "환율 리스크", text: "2024 평균 1,363원\n2025 평균 1,421원 고환율 지속", color: C.orange },
    ];

    trends.forEach((t, i) => {
      const y = 0.9 + i * 1.1;
      slide.addShape(pres.shapes.RECTANGLE, {
        x: 5.8, y, w: 3.8, h: 0.9,
        fill: { color: C.white }, shadow: makeCardShadow()
      });
      slide.addShape(pres.shapes.RECTANGLE, {
        x: 5.8, y, w: 0.06, h: 0.9,
        fill: { color: t.color }
      });
      slide.addText(t.title, {
        x: 6.1, y: y + 0.08, w: 3.3, h: 0.25,
        fontSize: 11, fontFace: "Calibri", color: t.color, bold: true, margin: 0
      });
      slide.addText(t.text, {
        x: 6.1, y: y + 0.35, w: 3.3, h: 0.45,
        fontSize: 10, fontFace: "Calibri", color: C.textDark, margin: 0, lineSpacingMultiple: 1.3
      });
    });

    // Bottom: SWOT Summary
    slide.addText("SWOT 분석 요약", {
      x: 0.5, y: 3.55, w: 3, h: 0.35,
      fontSize: 13, fontFace: "Calibri", color: C.textDark, bold: true, margin: 0
    });

    const swot = [
      { label: "S", title: "강점", items: "전기로 기반 탄소저감 우위\n해외 7개국 81만톤 글로벌 네트워크", color: C.accent },
      { label: "W", title: "약점", items: "원가율 90.3% 수익성 악화\n가동률 49.6%(생산/능력)", color: C.orange },
      { label: "O", title: "기회", items: "반덤핑 관세 → 중국산 수입 감소\nSeAH Wind 해상풍력 신성장", color: C.green },
      { label: "T", title: "위협", items: "건설투자 -8.8% 내수 위축\nHR Coil 가격 하락 추세", color: C.red },
    ];

    swot.forEach((s, i) => {
      const x = 0.5 + i * 2.35;
      slide.addShape(pres.shapes.RECTANGLE, {
        x, y: 3.9, w: 2.15, h: 1.3,
        fill: { color: C.white }, shadow: makeCardShadow()
      });
      // Header
      slide.addShape(pres.shapes.RECTANGLE, {
        x, y: 3.9, w: 2.15, h: 0.35,
        fill: { color: s.color }
      });
      slide.addText(`${s.label} · ${s.title}`, {
        x, y: 3.9, w: 2.15, h: 0.35,
        fontSize: 11, fontFace: "Calibri", color: C.white, bold: true,
        align: "center", valign: "middle"
      });
      slide.addText(s.items, {
        x: x + 0.12, y: 4.3, w: 1.9, h: 0.8,
        fontSize: 9, fontFace: "Calibri", color: C.textDark, margin: 0, lineSpacingMultiple: 1.4
      });
    });

    addFooter(slide, 8, TOTAL_PAGES);
  }

  // ============================================================
  // SLIDE 9: RESPONSE STRATEGY
  // ============================================================
  {
    const slide = pres.addSlide();
    slide.background = { color: C.offWhite };
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent }
    });

    slide.addText("전략적 대응 방안", {
      x: 0.5, y: 0.25, w: 7, h: 0.5,
      fontSize: 20, fontFace: "Calibri", color: C.textDark, bold: true, margin: 0
    });
    slide.addText("사업환경 변화에 따른 단기·중기 대응 전략", {
      x: 0.5, y: 0.7, w: 7, h: 0.3,
      fontSize: 11, fontFace: "Calibri", color: C.textMid, margin: 0
    });

    // Strategy cards — 2x2 grid
    const strategies = [
      {
        title: "원가 구조 개선 (원가율 90.3%→85%)",
        icon: icons.moneyDk,
        items: [
          "철스크랩 조달 다변화 (국내외 소싱 최적화, $330~370 변동 대응)",
          "에너지 효율 개선 (온실가스 64,078 tCO2-eq 감축 연계)",
          "가동률 제고 (현 49.6% → 생산능력 160만톤 활용 극대화)"
        ]
      },
      {
        title: "해외 사업 확대 (7개국 81만톤)",
        icon: icons.globeDk,
        items: [
          "SSUSA(미국) OCTG 20만톤 — IRA 수혜 확대",
          "SSV(베트남) ERW 24만톤 — 동남아 성장 거점",
          "Inox Tech(이탈리아) — 유럽 특수관 시장 공략"
        ]
      },
      {
        title: "신성장 동력 (SeAH Wind)",
        icon: icons.cogsDk,
        items: [
          "해상풍력 모노파일 사업 — RCPS 투자 1,479억원",
          "글로벌 해상풍력 시장 확대에 따른 선제 투자",
          "친환경 에너지 전환 트렌드와 시너지"
        ]
      },
      {
        title: "ESG·주주환원 강화",
        icon: icons.industryDk,
        items: [
          "배당성향 52%, 주당 5,500원 유지 → 주주 신뢰",
          "전기로 기반 저탄소 생산 우위 적극 활용",
          "R&D 38억원(매출 0.3%) → 고부가 제품 개발 확대"
        ]
      },
    ];

    strategies.forEach((s, i) => {
      const col = i % 2;
      const row = Math.floor(i / 2);
      const x = 0.5 + col * 4.7;
      const y = 1.15 + row * 2.1;
      const w = 4.4;
      const h = 1.9;

      // Card
      slide.addShape(pres.shapes.RECTANGLE, {
        x, y, w, h, fill: { color: C.white }, shadow: makeCardShadow()
      });
      // Top accent
      slide.addShape(pres.shapes.RECTANGLE, {
        x, y, w, h: 0.05, fill: { color: C.accent }
      });

      // Icon + Title
      slide.addImage({ data: s.icon, x: x + 0.2, y: y + 0.2, w: 0.32, h: 0.32 });
      slide.addText(s.title, {
        x: x + 0.6, y: y + 0.15, w: 3.5, h: 0.4,
        fontSize: 14, fontFace: "Calibri", color: C.textDark, bold: true, margin: 0, valign: "middle"
      });

      // Items
      const itemTexts = s.items.map((item, j) => ({
        text: item,
        options: { bullet: true, fontSize: 10, color: C.textDark, breakLine: j < s.items.length - 1 }
      }));
      slide.addText(itemTexts, {
        x: x + 0.25, y: y + 0.6, w: w - 0.5, h: 1.2,
        fontFace: "Calibri", margin: 0, paraSpaceAfter: 4
      });
    });

    addFooter(slide, 9, TOTAL_PAGES);
  }

  // ============================================================
  // SLIDE 10: MID/LONG-TERM PLAN
  // ============================================================
  {
    const slide = pres.addSlide();
    slide.background = { color: C.offWhite };
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent }
    });

    slide.addText("03  중장기 경영계획 수립 (2026~2028)", {
      x: 0.5, y: 0.25, w: 8, h: 0.5,
      fontSize: 20, fontFace: "Calibri", color: C.textDark, bold: true, margin: 0
    });

    // Vision statement
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y: 0.9, w: 9, h: 0.55,
      fill: { color: C.navy }
    });
    slide.addText("비전:  글로벌 철강·에너지 소재 기업으로의 도약", {
      x: 0.5, y: 0.9, w: 9, h: 0.55,
      fontSize: 15, fontFace: "Calibri", color: C.white, bold: true,
      align: "center", valign: "middle"
    });

    // 3-year roadmap timeline
    const years = [
      {
        year: "2026", title: "수익성 회복", color: C.steelBlue,
        items: "원가율 85% 이하 목표\nSeAH Wind 본격 가동\n가동률 60% 이상 회복"
      },
      {
        year: "2027", title: "성장 가속", color: C.accent,
        items: "해외법인 매출 비중 확대\n고부가 제품 믹스 강화\nR&D 투자 확대"
      },
      {
        year: "2028", title: "글로벌 도약", color: C.navy,
        items: "글로벌 매출 비중 30%+\n해상풍력 수익 본격화\n매출 목표: 1.8조원"
      },
    ];

    years.forEach((yr, i) => {
      const x = 0.5 + i * 3.15;

      // Arrow connector (except last)
      if (i < 2) {
        slide.addShape(pres.shapes.LINE, {
          x: x + 2.9, y: 2.3, w: 0.3, h: 0,
          line: { color: C.midGray, width: 1.5, dashType: "dash" }
        });
      }

      // Year card
      slide.addShape(pres.shapes.RECTANGLE, {
        x, y: 1.7, w: 2.9, h: 2.0,
        fill: { color: C.white }, shadow: makeCardShadow()
      });
      slide.addShape(pres.shapes.RECTANGLE, {
        x, y: 1.7, w: 2.9, h: 0.05,
        fill: { color: yr.color }
      });

      // Year badge
      slide.addShape(pres.shapes.RECTANGLE, {
        x: x + 0.15, y: 1.85, w: 0.8, h: 0.35,
        fill: { color: yr.color }
      });
      slide.addText(yr.year, {
        x: x + 0.15, y: 1.85, w: 0.8, h: 0.35,
        fontSize: 12, fontFace: "Calibri", color: C.white, bold: true,
        align: "center", valign: "middle"
      });

      slide.addText(yr.title, {
        x: x + 1.05, y: 1.85, w: 1.7, h: 0.35,
        fontSize: 13, fontFace: "Calibri", color: yr.color, bold: true, margin: 0, valign: "middle"
      });

      slide.addText(yr.items, {
        x: x + 0.2, y: 2.35, w: 2.5, h: 1.2,
        fontSize: 10, fontFace: "Calibri", color: C.textDark, margin: 0, lineSpacingMultiple: 1.5
      });
    });

    // KPI targets table
    slide.addText("핵심 KPI 목표", {
      x: 0.5, y: 3.9, w: 3, h: 0.35,
      fontSize: 13, fontFace: "Calibri", color: C.textDark, bold: true, margin: 0
    });

    const kpiTable = [
      [
        { text: "KPI", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 10, align: "center" } },
        { text: "2025 실적", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 10, align: "center" } },
        { text: "2026 목표", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 10, align: "center" } },
        { text: "2027 목표", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 10, align: "center" } },
        { text: "2028 목표", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 10, align: "center" } },
      ],
      [
        { text: "매출액(조원)", options: { fontSize: 10, fill: { color: C.white }, align: "center" } },
        { text: "1.48", options: { fontSize: 10, fill: { color: C.white }, align: "center" } },
        { text: "1.55", options: { fontSize: 10, fill: { color: C.white }, align: "center" } },
        { text: "1.65", options: { fontSize: 10, fill: { color: C.white }, align: "center" } },
        { text: "1.80", options: { fontSize: 10, fill: { color: C.white }, align: "center", bold: true, color: C.accent } },
      ],
      [
        { text: "영업이익률(%)", options: { fontSize: 10, fill: { color: C.paleBlue }, align: "center" } },
        { text: "3.3", options: { fontSize: 10, fill: { color: C.paleBlue }, align: "center", color: C.red } },
        { text: "6.0", options: { fontSize: 10, fill: { color: C.paleBlue }, align: "center" } },
        { text: "8.0", options: { fontSize: 10, fill: { color: C.paleBlue }, align: "center" } },
        { text: "10.0", options: { fontSize: 10, fill: { color: C.paleBlue }, align: "center", bold: true, color: C.accent } },
      ],
      [
        { text: "ROE(%)", options: { fontSize: 10, fill: { color: C.white }, align: "center" } },
        { text: "2.7", options: { fontSize: 10, fill: { color: C.white }, align: "center", color: C.red } },
        { text: "5.0", options: { fontSize: 10, fill: { color: C.white }, align: "center" } },
        { text: "7.0", options: { fontSize: 10, fill: { color: C.white }, align: "center" } },
        { text: "10.0", options: { fontSize: 10, fill: { color: C.white }, align: "center", bold: true, color: C.accent } },
      ],
    ];

    slide.addTable(kpiTable, {
      x: 0.5, y: 4.2, w: 9, h: 1.1,
      border: { pt: 0.5, color: C.lightGray },
      colW: [2.2, 1.7, 1.7, 1.7, 1.7],
      fontFace: "Calibri",
    });

    addFooter(slide, 10, TOTAL_PAGES);
  }

  // ============================================================
  // SLIDE 13: BUDGET & MEETING MANAGEMENT
  // ============================================================
  {
    const slide = pres.addSlide();
    slide.background = { color: C.offWhite };
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent }
    });

    slide.addText("04  투자예산 및 회의체 운영관리", {
      x: 0.5, y: 0.25, w: 7, h: 0.5,
      fontSize: 20, fontFace: "Calibri", color: C.textDark, bold: true, margin: 0
    });

    // Budget execution chart
    slide.addText("부문별 예산 집행 현황 (2025년 11월 누계)", {
      x: 0.5, y: 0.85, w: 5, h: 0.3,
      fontSize: 12, fontFace: "Calibri", color: C.textMid, margin: 0
    });

    slide.addChart(pres.charts.BAR, [
      {
        name: "예산(억원)",
        labels: ["설비투자", "R&D", "해외사업", "IT/디지털", "인력개발"],
        values: [500, 45, 200, 60, 40]
      },
      {
        name: "집행(억원)",
        labels: ["설비투자", "R&D", "해외사업", "IT/디지털", "인력개발"],
        values: [420, 38, 165, 52, 35]
      }
    ], {
      x: 0.3, y: 1.1, w: 5.0, h: 2.3, barDir: "col",
      chartColors: [C.steelBlue, C.accent],
      chartArea: { fill: { color: C.white }, roundedCorners: true },
      catAxisLabelColor: C.textMid, catAxisLabelFontSize: 8,
      valAxisLabelColor: C.textMid, valAxisLabelFontSize: 8,
      valGridLine: { color: C.lightGray, size: 0.5 },
      catGridLine: { style: "none" },
      showLegend: true, legendPos: "b", legendFontSize: 9,
      showValue: true, dataLabelPosition: "outEnd", dataLabelColor: C.textDark, dataLabelFontSize: 8,
    });

    // Execution rate cards
    slide.addText("집행률 현황", {
      x: 5.6, y: 0.85, w: 4, h: 0.3,
      fontSize: 12, fontFace: "Calibri", color: C.textMid, margin: 0
    });

    const execRates = [
      { dept: "설비투자", rate: "84.0%", status: "양호" },
      { dept: "R&D", rate: "84.4%", status: "양호" },
      { dept: "해외사업", rate: "82.5%", status: "양호" },
      { dept: "IT/디지털", rate: "86.7%", status: "정상" },
      { dept: "인력개발", rate: "87.5%", status: "정상" },
    ];

    execRates.forEach((item, i) => {
      const y = 1.2 + i * 0.42;
      slide.addShape(pres.shapes.RECTANGLE, {
        x: 5.6, y, w: 4.0, h: 0.35,
        fill: { color: i % 2 === 0 ? C.white : C.paleBlue }
      });
      slide.addText(item.dept, {
        x: 5.75, y, w: 1.5, h: 0.35,
        fontSize: 10, fontFace: "Calibri", color: C.textDark, valign: "middle", margin: 0
      });
      slide.addText(item.rate, {
        x: 7.3, y, w: 1.0, h: 0.35,
        fontSize: 11, fontFace: "Calibri", color: C.accent, bold: true, valign: "middle", margin: 0, align: "center"
      });
      slide.addText(item.status, {
        x: 8.4, y, w: 1.0, h: 0.35,
        fontSize: 10, fontFace: "Calibri", color: C.green, valign: "middle", margin: 0, align: "center"
      });
    });

    // Meeting management
    slide.addText("주요 회의체 운영 현황", {
      x: 0.5, y: 3.6, w: 5, h: 0.35,
      fontSize: 13, fontFace: "Calibri", color: C.textDark, bold: true, margin: 0
    });

    const meetingTable = [
      [
        { text: "회의체", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 10, align: "center" } },
        { text: "주기", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 10, align: "center" } },
        { text: "참석자", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 10, align: "center" } },
        { text: "주요 안건", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 10, align: "center" } },
        { text: "의사결정 사항", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 10, align: "center" } },
      ],
      [
        { text: "경영전략회의", options: { fontSize: 9, fill: { color: C.white }, align: "center" } },
        { text: "월 1회", options: { fontSize: 9, fill: { color: C.white }, align: "center" } },
        { text: "CEO, C-Level", options: { fontSize: 9, fill: { color: C.white }, align: "center" } },
        { text: "경영실적 리뷰", options: { fontSize: 9, fill: { color: C.white } } },
        { text: "전략 방향 승인", options: { fontSize: 9, fill: { color: C.white } } },
      ],
      [
        { text: "투자심의위원회", options: { fontSize: 9, fill: { color: C.paleBlue }, align: "center" } },
        { text: "분기 1회", options: { fontSize: 9, fill: { color: C.paleBlue }, align: "center" } },
        { text: "CFO, 사업부장", options: { fontSize: 9, fill: { color: C.paleBlue }, align: "center" } },
        { text: "투자안 심의", options: { fontSize: 9, fill: { color: C.paleBlue } } },
        { text: "투자 승인/보류", options: { fontSize: 9, fill: { color: C.paleBlue } } },
      ],
      [
        { text: "실적점검회의", options: { fontSize: 9, fill: { color: C.white }, align: "center" } },
        { text: "주 1회", options: { fontSize: 9, fill: { color: C.white }, align: "center" } },
        { text: "팀장급 이상", options: { fontSize: 9, fill: { color: C.white }, align: "center" } },
        { text: "주간 실적 점검", options: { fontSize: 9, fill: { color: C.white } } },
        { text: "실행과제 배분", options: { fontSize: 9, fill: { color: C.white } } },
      ],
    ];

    slide.addTable(meetingTable, {
      x: 0.5, y: 3.95, w: 9, h: 1.2,
      border: { pt: 0.5, color: C.lightGray },
      colW: [1.8, 1.2, 1.6, 2.2, 2.2],
      fontFace: "Calibri",
    });

    addFooter(slide, 11, TOTAL_PAGES);
  }

  // ============================================================
  // SLIDE 14: CLOSING
  // ============================================================
  {
    const slide = pres.addSlide();
    slide.background = { color: C.darkNavy };

    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent }
    });

    // Decorative circles
    slide.addShape(pres.shapes.OVAL, {
      x: -0.5, y: 3.5, w: 3, h: 3,
      fill: { color: C.navy, transparency: 50 }
    });
    slide.addShape(pres.shapes.OVAL, {
      x: 8, y: -0.5, w: 2.5, h: 2.5,
      fill: { color: C.steelBlue, transparency: 60 }
    });

    slide.addText("감사합니다", {
      x: 1, y: 1.5, w: 8, h: 1.0,
      fontSize: 40, fontFace: "Calibri", color: C.white, bold: true,
      align: "center", valign: "middle"
    });

    slide.addText("데이터 기반 분석 역량으로\n경영기획의 핵심 가치를 실현하겠습니다.", {
      x: 1.5, y: 2.5, w: 7, h: 0.9,
      fontSize: 16, fontFace: "Calibri", color: C.lightBlue,
      align: "center", lineSpacingMultiple: 1.5
    });

    // Key competency badges
    const badges = ["데이터 분석", "보고서 작성", "전략적 사고", "문제 해결"];
    badges.forEach((badge, i) => {
      const x = 1.8 + i * 1.7;
      slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x, y: 3.7, w: 1.5, h: 0.4,
        fill: { color: C.accent, transparency: 20 }, rectRadius: 0.08,
        line: { color: C.accent, width: 1 }
      });
      slide.addText(badge, {
        x, y: 3.7, w: 1.5, h: 0.4,
        fontSize: 10, fontFace: "Calibri", color: C.lightBlue, bold: true,
        align: "center", valign: "middle"
      });
    });

    slide.addText("세아제강 경영기획 직무역량 포트폴리오  |  2026년 3월", {
      x: 1, y: 4.6, w: 8, h: 0.4,
      fontSize: 10, fontFace: "Calibri", color: C.midGray, align: "center"
    });
  }

  // ============================================================
  // SAVE
  // ============================================================
  await pres.writeFile({ fileName: "C:\\Users\\pc\\Desktop\\데이터분석\\세아제강\\files\\세아제강_경영기획_포트폴리오.pptx" });
  console.log("Portfolio created successfully!");
}

main().catch(err => { console.error(err); process.exit(1); });
