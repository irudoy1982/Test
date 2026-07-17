import fs from "node:fs/promises";
import path from "node:path";
import { Presentation, PresentationFile } from "@oai/artifact-tool";

const SCRIPT_DIR = path.dirname(new URL(import.meta.url).pathname.replace(/^\/(.:)/, "$1"));
const ROOT = process.env.AUDIT_PROJECT_ROOT
  ? path.resolve(process.env.AUDIT_PROJECT_ROOT)
  : path.resolve(SCRIPT_DIR, "..");
const STATIC = path.join(ROOT, "static");
const OUT = process.argv[2] ? path.resolve(process.argv[2]) : STATIC;

const W = 1280;
const H = 720;

async function readImageBlob(imagePath) {
  const bytes = await fs.readFile(imagePath);
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}

function addRect(slide, x, y, w, h, fill, radius = "rounded-lg", line = "none") {
  return slide.shapes.add({
    geometry: radius === "none" ? "rect" : "roundRect",
    position: { left: x, top: y, width: w, height: h },
    fill,
    line: line === "none" ? { style: "solid", fill: "none", width: 0 } : line,
    ...(radius === "none" ? {} : { borderRadius: radius }),
  });
}

function addText(slide, text, x, y, w, h, style = {}) {
  const shape = slide.shapes.add({
    geometry: "textbox",
    position: { left: x, top: y, width: w, height: h },
    fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  shape.text = text;
  shape.text.style = {
    fontFamily: "Arial",
    fontSize: 18,
    color: "#182033",
    ...style,
  };
  return shape;
}

async function addImage(slide, imagePath, position, fit = "contain", radius = "none") {
  const contentType = imagePath.toLowerCase().endsWith(".jpg") || imagePath.toLowerCase().endsWith(".jpeg")
    ? "image/jpeg"
    : "image/png";
  slide.images.add({
    blob: await readImageBlob(imagePath),
    contentType,
    alt: path.basename(imagePath),
    fit,
    position,
    geometry: radius === "none" ? "rect" : "roundRect",
    ...(radius === "none" ? {} : { borderRadius: radius }),
  });
}

function addFooter(slide, brand, page, colors) {
  addRect(slide, 0, 696, W, 24, colors.accent, "none");
  addText(slide, `${brand} Audit System  |  by Ivan Rudoy`, 42, 699, 600, 16, {
    fontSize: 10,
    bold: true,
    color: "#FFFFFF",
  });
  addText(slide, String(page).padStart(2, "0"), 1190, 699, 45, 16, {
    fontSize: 10,
    bold: true,
    color: "#FFFFFF",
    alignment: "right",
  });
}

async function addHeader(slide, cfg, title, page) {
  addRect(slide, 0, 0, 18, H, cfg.colors.accent, "none");
  await addImage(slide, cfg.logo, { left: 1030, top: 26, width: 190, height: 60 }, "contain");
  addText(slide, title, 58, 44, 870, 56, {
    fontSize: 34,
    bold: true,
    color: cfg.colors.dark,
  });
  addRect(slide, 58, 112, 104, 7, cfg.colors.accent, "rounded-lg");
  addFooter(slide, cfg.brand, page, cfg.colors);
}

function addBulletRows(slide, items, cfg, yStart, rowH = 112) {
  items.forEach((item, index) => {
    const y = yStart + index * rowH;
    addText(slide, String(index + 1).padStart(2, "0"), 64, y + 6, 48, 28, {
      fontSize: 17,
      bold: true,
      color: cfg.colors.accent,
    });
    addText(slide, item, 126, y, 1055, rowH - 14, {
      fontSize: 18,
      color: cfg.colors.dark,
    });
    if (index < items.length - 1) addRect(slide, 126, y + rowH - 18, 1055, 1, cfg.colors.rule, "none");
  });
}

function addRecommendationCards(slide, startIndex, cfg) {
  const positions = [
    { x: 58, y: 150 },
    { x: 638, y: 150 },
    { x: 58, y: 405 },
    { x: 638, y: 405 },
  ];
  for (let index = 0; index < 4; index += 1) {
    const number = startIndex + index;
    const { x, y } = positions[index];
    addRect(slide, x, y, 542, 225, "#FFFFFF", "rounded-lg", {
      style: "solid",
      fill: cfg.colors.rule,
      width: 1,
    });
    addText(slide, String(number).padStart(2, "0"), x + 20, y + 18, 38, 26, {
      fontSize: 16,
      bold: true,
      color: cfg.colors.accent,
    });
    addText(slide, `{{REC_${number}_TITLE}}`, x + 64, y + 15, 454, 42, {
      fontSize: 18,
      bold: true,
      color: cfg.colors.dark,
    });
    addText(slide, `{{REC_${number}_ACTION}}`, x + 22, y + 65, 498, 58, {
      fontSize: 15,
      color: cfg.colors.muted,
    });
    addRect(slide, x + 22, y + 132, 498, 1, cfg.colors.rule, "none");
    addText(slide, "Решение", x + 22, y + 146, 228, 18, {
      fontSize: 12,
      bold: true,
      color: cfg.colors.accent,
    });
    addText(slide, `{{REC_${number}_SOLUTION}}`, x + 22, y + 166, 238, 43, {
      fontSize: 14,
      bold: true,
      color: cfg.colors.dark,
    });
    addText(slide, "Производители", x + 280, y + 146, 238, 18, {
      fontSize: 12,
      bold: true,
      color: cfg.colors.accent,
    });
    addText(slide, `{{REC_${number}_VENDORS}}`, x + 280, y + 166, 238, 43, {
      fontSize: 14,
      bold: true,
      color: cfg.colors.dark,
    });
  }
}

async function buildTemplate(cfg) {
  const deck = Presentation.create({ slideSize: { width: W, height: H } });

  // 1. Cover
  {
    const slide = deck.slides.add();
    slide.background.fill = "#FFFFFF";
    addRect(slide, 0, 0, 24, H, cfg.colors.accent, "none");
    addRect(slide, 24, 0, 742, H, cfg.colors.soft, "none");
    await addImage(slide, cfg.logo, { left: 72, top: 55, width: 250, height: 90 }, "contain");
    addText(slide, "ЭКСПЕРТНЫЙ АУДИТ ИТ И ИБ", 72, 185, 610, 30, {
      fontSize: 17,
      bold: true,
      color: cfg.colors.accent,
    });
    addText(slide, "{{COMPANY}}", 72, 235, 630, 145, {
      fontSize: 50,
      bold: true,
      color: cfg.colors.dark,
    });
    addText(slide, "{{INDUSTRY}}  •  {{CITY}}", 72, 392, 620, 42, {
      fontSize: 22,
      color: cfg.colors.muted,
    });
    addRect(slide, 72, 495, 190, 104, cfg.colors.dark, "rounded-xl");
    addText(slide, "Зрелость ИТ", 88, 512, 158, 25, {
      fontSize: 15,
      bold: true,
      color: "#FFFFFF",
      alignment: "center",
    });
    addText(slide, "{{IT_SCORE}}%", 88, 541, 158, 44, {
      fontSize: 34,
      bold: true,
      color: "#FFFFFF",
      alignment: "center",
    });
    addRect(slide, 278, 495, 190, 104, cfg.colors.accent, "rounded-xl");
    addText(slide, "Зрелость ИБ", 294, 512, 158, 25, {
      fontSize: 15,
      bold: true,
      color: "#FFFFFF",
      alignment: "center",
    });
    addText(slide, "{{SCORE}}%", 294, 541, 158, 44, {
      fontSize: 34,
      bold: true,
      color: "#FFFFFF",
      alignment: "center",
    });
    addText(slide, "{{DATE}}", 505, 542, 180, 34, {
      fontSize: 18,
      bold: true,
      color: cfg.colors.dark,
    });
    await addImage(slide, cfg.cover, { left: 766, top: 0, width: 514, height: 720 }, "cover");
    addRect(slide, 766, 0, 12, H, cfg.colors.accent, "none");
    addText(slide, "by Ivan Rudoy", 72, 650, 220, 24, {
      fontSize: 14,
      bold: true,
      color: cfg.colors.muted,
    });
  }

  // 2. Company profile
  {
    const slide = deck.slides.add();
    slide.background.fill = "#FFFFFF";
    await addHeader(slide, cfg, "Кто стоит за аудитом", 2);

    addText(slide, cfg.companyName, 58, 150, 520, 46, {
      fontSize: 25,
      bold: true,
      color: cfg.colors.accent,
    });
    addText(slide, cfg.companySummary, 58, 205, 520, 92, {
      fontSize: 18,
      color: cfg.colors.dark,
    });
    addText(slide, "От аудита и проектирования - до внедрения, поддержки и развития.", 58, 322, 520, 62, {
      fontSize: 20,
      bold: true,
      color: cfg.colors.dark,
    });

    const stats = [
      [cfg.foundedYear, "год основания"],
      ["70+", "крупных проектов"],
      ["40+", "технологических партнеров"],
    ];
    stats.forEach(([value, label], i) => {
      const x = 58 + i * 178;
      addText(slide, value, x, 432, 148, 46, {
        fontSize: i === 1 ? 36 : 31,
        bold: true,
        color: cfg.colors.dark,
      });
      addText(slide, label, x, 481, 148, 62, {
        fontSize: 13,
        color: cfg.colors.muted,
      });
    });
    addRect(slide, 58, 558, 520, 58, cfg.colors.soft, "rounded-lg", {
      style: "solid",
      fill: cfg.colors.rule,
      width: 1,
    });
    addText(slide, "Одна команда отвечает за архитектуру, безопасность и практическую реализацию.", 78, 572, 480, 30, {
      fontSize: 15,
      bold: true,
      color: cfg.colors.dark,
      alignment: "center",
    });

    const directions = [
      ["01", "Информационная безопасность", "Аудит, проектирование защиты, внедрение, поддержка и обучение."],
      ["02", "Системная интеграция", "Сетевая, серверная и вычислительная инфраструктура под ключ."],
      ["03", "Автоматизация процессов", "СЭД, low-code, ESM/BPM и цифровые решения для бизнеса."],
      ["04", "Сопровождение и развитие", "Рабочие сессии, внедрение, поддержка и передача знаний команде."],
    ];
    directions.forEach(([number, title, text], i) => {
      const y = 150 + i * 119;
      addText(slide, number, 632, y + 4, 44, 28, {
        fontSize: 16,
        bold: true,
        color: cfg.colors.accent,
      });
      addText(slide, title, 690, y, 486, 30, {
        fontSize: 19,
        bold: true,
        color: cfg.colors.dark,
      });
      addText(slide, text, 690, y + 36, 486, 52, {
        fontSize: 15,
        color: cfg.colors.muted,
      });
      if (i < directions.length - 1) addRect(slide, 690, y + 103, 486, 1, cfg.colors.rule, "none");
    });
  }

  // 3. Executive summary
  {
    const slide = deck.slides.add();
    slide.background.fill = "#FFFFFF";
    await addHeader(slide, cfg, "Главный вывод аудита", 3);
    addText(slide, "{{SUMMARY_TITLE}}", 58, 146, 1118, 72, {
      fontSize: 28,
      bold: true,
      color: cfg.colors.accent,
    });
    addBulletRows(slide, ["{{SUMMARY_1}}", "{{SUMMARY_2}}", "{{SUMMARY_3}}", "{{SUMMARY_4}}"], cfg, 246, 100);
  }

  // 4. Infrastructure profile
  {
    const slide = deck.slides.add();
    slide.background.fill = "#FFFFFF";
    await addHeader(slide, cfg, "Масштаб инфраструктуры задает три приоритета", 4);
    const metrics = [
      ["{{USERS}}", "рабочих мест"],
      ["{{SERVERS}}", "серверов"],
      ["{{PUBLIC}}", "публичные сервисы"],
      ["{{BUSINESS}}", "бизнес-системы"],
    ];
    metrics.forEach(([value, label], i) => {
      const x = 58 + i * 288;
      addRect(slide, x, 155, 260, 126, i === 0 ? cfg.colors.accent : cfg.colors.soft, "rounded-lg", {
        style: "solid",
        fill: i === 0 ? cfg.colors.accent : cfg.colors.rule,
        width: 1,
      });
      addText(slide, value, x + 18, 174, 224, 48, {
        fontSize: 32,
        bold: true,
        color: i === 0 ? "#FFFFFF" : cfg.colors.dark,
        alignment: "center",
      });
      addText(slide, label, x + 18, 229, 224, 28, {
        fontSize: 15,
        color: i === 0 ? "#FFFFFF" : cfg.colors.muted,
        alignment: "center",
      });
    });
    addText(slide, "Профиль", 58, 316, 180, 26, { fontSize: 17, bold: true, color: cfg.colors.accent });
    addText(slide, "{{PROFILE}}", 58, 348, 1118, 62, { fontSize: 18, bold: true, color: cfg.colors.dark });
    addText(slide, "Управленческий фокус", 58, 432, 300, 28, { fontSize: 18, bold: true, color: cfg.colors.accent });
    for (let i = 0; i < 3; i += 1) {
      const x = 58 + i * 382;
      addRect(slide, x, 482, 7, 140, cfg.colors.accent, "rounded-lg");
      addText(slide, `{{FOCUS_${i + 1}_TITLE}}`, x + 24, 478, 326, 38, {
        fontSize: 18,
        bold: true,
        color: cfg.colors.dark,
      });
      addText(slide, `{{FOCUS_${i + 1}_TEXT}}`, x + 24, 525, 326, 98, {
        fontSize: 15,
        color: cfg.colors.muted,
      });
    }
  }

  // 5. Risks
  {
    const slide = deck.slides.add();
    slide.background.fill = "#FFFFFF";
    await addHeader(slide, cfg, "Четыре риска требуют управленческого внимания", 5);
    const severityColors = ["#B42318", "#D9480F", "#E67700", cfg.colors.accent];
    for (let i = 0; i < 4; i += 1) {
      const y = 146 + i * 128;
      addRect(slide, 58, y, 110, 106, severityColors[i], "rounded-lg");
      addText(slide, `{{RISK_${i + 1}_LEVEL}}`, 68, y + 37, 90, 28, {
        fontSize: 13,
        bold: true,
        color: "#FFFFFF",
        alignment: "center",
      });
      addText(slide, `{{RISK_${i + 1}_TITLE}}`, 196, y, 980, 31, {
        fontSize: 18,
        bold: true,
        color: cfg.colors.dark,
      });
      addText(slide, `{{RISK_${i + 1}_IMPACT}}`, 196, y + 32, 980, 35, {
        fontSize: 14,
        color: cfg.colors.muted,
      });
      addText(slide, "Рекомендуем", 196, y + 76, 104, 20, {
        fontSize: 13,
        bold: true,
        color: cfg.colors.accent,
      });
      addText(slide, `{{RISK_${i + 1}_RECOMMENDATION}}`, 304, y + 72, 872, 34, {
        fontSize: 14,
        bold: true,
        color: cfg.colors.dark,
      });
      if (i < 3) addRect(slide, 196, y + 116, 980, 1, cfg.colors.rule, "none");
    }
  }

  // 6. Priority recommendations
  {
    const slide = deck.slides.add();
    slide.background.fill = "#FFFFFF";
    await addHeader(slide, cfg, "Рекомендации: первоочередные меры", 6);
    addRecommendationCards(slide, 1, cfg);
  }

  // 7. Next recommendations
  {
    const slide = deck.slides.add();
    slide.background.fill = "#FFFFFF";
    await addHeader(slide, cfg, "Рекомендации: следующий этап", 7);
    addRecommendationCards(slide, 5, cfg);
  }

  // 8. Roadmap
  {
    const slide = deck.slides.add();
    slide.background.fill = "#FFFFFF";
    await addHeader(slide, cfg, "План на 90 дней переводит выводы в действия", 8);
    const phases = [
      ["0–30 дней", "{{ROADMAP_1_1}}", "{{ROADMAP_1_2}}"],
      ["31–60 дней", "{{ROADMAP_2_1}}", "{{ROADMAP_2_2}}"],
      ["61–90 дней", "{{ROADMAP_3_1}}", "{{ROADMAP_3_2}}"],
    ];
    phases.forEach(([phase, one, two], i) => {
      const x = 58 + i * 382;
      addRect(slide, x, 158, 354, 454, i === 0 ? cfg.colors.soft : "#FFFFFF", "rounded-lg", {
        style: "solid",
        fill: cfg.colors.rule,
        width: 1,
      });
      addRect(slide, x, 158, 354, 68, i === 0 ? cfg.colors.accent : cfg.colors.dark, "rounded-lg");
      addText(slide, phase, x + 24, 178, 306, 30, {
        fontSize: 21,
        bold: true,
        color: "#FFFFFF",
        alignment: "center",
      });
      addText(slide, "01", x + 24, 263, 40, 28, { fontSize: 16, bold: true, color: cfg.colors.accent });
      addText(slide, one, x + 72, 255, 250, 118, { fontSize: 17, color: cfg.colors.dark });
      addRect(slide, x + 24, 389, 306, 1, cfg.colors.rule, "none");
      addText(slide, "02", x + 24, 424, 40, 28, { fontSize: 16, bold: true, color: cfg.colors.accent });
      addText(slide, two, x + 72, 416, 250, 134, { fontSize: 17, color: cfg.colors.dark });
    });
  }

  // 9. Decisions
  {
    const slide = deck.slides.add();
    slide.background.fill = "#FFFFFF";
    await addHeader(slide, cfg, "Зафиксируйте решения и следующий шаг", 9);
    const decisions = ["{{DECISION_1}}", "{{DECISION_2}}", "{{DECISION_3}}", "{{DECISION_4}}"];
    decisions.forEach((item, i) => {
      const y = 164 + i * 112;
      addText(slide, String(i + 1).padStart(2, "0"), 58, y, 48, 30, {
        fontSize: 16,
        bold: true,
        color: cfg.colors.accent,
      });
      addText(slide, item, 116, y - 4, 674, 78, {
        fontSize: 17,
        color: cfg.colors.dark,
      });
      if (i < 3) addRect(slide, 116, y + 88, 674, 1, cfg.colors.rule, "none");
    });

    addRect(slide, 830, 146, 364, 494, cfg.colors.soft, "rounded-xl", {
      style: "solid",
      fill: cfg.colors.rule,
      width: 1,
    });
    addText(slide, "Хотите получить больше сведений?", 862, 174, 300, 66, {
      fontSize: 25,
      bold: true,
      color: cfg.colors.dark,
      alignment: "center",
    });
    addText(slide, "Обсудим приоритеты, варианты решений и практический план внедрения.", 866, 242, 292, 60, {
      fontSize: 16,
      color: cfg.colors.muted,
      alignment: "center",
    });
    await addImage(slide, cfg.qr, { left: 914, top: 314, width: 196, height: 196 }, "contain");
    addText(slide, "Сканируйте QR, чтобы сохранить контакты", 860, 520, 304, 28, {
      fontSize: 13,
      color: cfg.colors.muted,
      alignment: "center",
    });
    addText(slide, cfg.email, 854, 558, 316, 24, {
      fontSize: 15,
      bold: true,
      color: cfg.colors.accent,
      alignment: "center",
    });
    addText(slide, cfg.phone, 854, 586, 316, 24, {
      fontSize: 15,
      bold: true,
      color: cfg.colors.dark,
      alignment: "center",
    });
  }

  return deck;
}

async function main() {
  await fs.mkdir(OUT, { recursive: true });
  const configs = [
    {
      key: "khalil",
      brand: "Khalil",
      foundedYear: "2020",
      companyName: "ТОО «Khalil Trade»",
      companySummary: "Системный интегратор и поставщик ИТ-решений для государственных и корпоративных организаций.",
      logo: path.join(STATIC, "presentation_khalil_logo.png"),
      cover: path.join(STATIC, "presentation_khalil_cover.png"),
      qr: path.join(STATIC, "presentation_khalil_qr.png"),
      email: "info@khalilgroup.kz",
      phone: "+7 706 701 48 35",
      colors: {
        accent: "#FF6412",
        dark: "#161616",
        soft: "#FFF3EB",
        muted: "#667085",
        rule: "#E4E7EC",
      },
    },
    {
      key: "btg",
      brand: "BTG",
      foundedYear: "2019",
      companyName: "ТОО «Bolashak Tamer Group»",
      companySummary: "Системный интегратор и поставщик ИТ-решений для государственных и корпоративных организаций.",
      logo: path.join(STATIC, "presentation_btg_logo.png"),
      cover: path.join(STATIC, "presentation_btg_cover.png"),
      qr: path.join(STATIC, "presentation_btg_qr.png"),
      email: "info@btgroup.kz",
      phone: "+7 706 700 48 35",
      colors: {
        accent: "#2048A8",
        dark: "#102A63",
        soft: "#EEF3FF",
        muted: "#596780",
        rule: "#D7E0F2",
      },
    },
  ];

  for (const cfg of configs) {
    const deck = await buildTemplate(cfg);
    const pptx = await PresentationFile.exportPptx(deck);
    const outputPath = path.join(OUT, `audit_presentation_${cfg.key}.pptx`);
    await pptx.save(outputPath);
    await fs.rm(`${outputPath}.inspect.ndjson`, { force: true });
  }
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
