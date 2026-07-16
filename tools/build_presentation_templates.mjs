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
    addRect(slide, 72, 495, 206, 104, cfg.colors.accent, "rounded-xl");
    addText(slide, "Зрелость ИБ", 92, 512, 166, 25, {
      fontSize: 15,
      bold: true,
      color: "#FFFFFF",
      alignment: "center",
    });
    addText(slide, "{{SCORE}}%", 92, 541, 166, 44, {
      fontSize: 34,
      bold: true,
      color: "#FFFFFF",
      alignment: "center",
    });
    addText(slide, "{{DATE}}", 316, 542, 210, 34, {
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

  // 2. Executive summary
  {
    const slide = deck.slides.add();
    slide.background.fill = "#FFFFFF";
    await addHeader(slide, cfg, "Главный вывод аудита", 2);
    addText(slide, "{{SUMMARY_TITLE}}", 58, 146, 1118, 72, {
      fontSize: 28,
      bold: true,
      color: cfg.colors.accent,
    });
    addBulletRows(slide, ["{{SUMMARY_1}}", "{{SUMMARY_2}}", "{{SUMMARY_3}}", "{{SUMMARY_4}}"], cfg, 246, 100);
  }

  // 3. Infrastructure profile
  {
    const slide = deck.slides.add();
    slide.background.fill = "#FFFFFF";
    await addHeader(slide, cfg, "Контекст инфраструктуры определяет приоритеты", 3);
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
    addText(slide, "Профиль", 58, 330, 180, 30, { fontSize: 18, bold: true, color: cfg.colors.accent });
    addText(slide, "{{PROFILE}}", 58, 372, 1118, 82, { fontSize: 20, bold: true, color: cfg.colors.dark });
    addText(slide, "Что это означает", 58, 490, 240, 30, { fontSize: 18, bold: true, color: cfg.colors.accent });
    addText(slide, "{{PROFILE_IMPLICATION}}", 58, 532, 1118, 104, { fontSize: 19, color: cfg.colors.dark });
  }

  // 4. Risks
  {
    const slide = deck.slides.add();
    slide.background.fill = "#FFFFFF";
    await addHeader(slide, cfg, "Четыре риска требуют управленческого внимания", 4);
    const severityColors = ["#B42318", "#D9480F", "#E67700", cfg.colors.accent];
    for (let i = 0; i < 4; i += 1) {
      const y = 148 + i * 126;
      addRect(slide, 58, y, 110, 94, severityColors[i], "rounded-lg");
      addText(slide, `{{RISK_${i + 1}_LEVEL}}`, 68, y + 31, 90, 28, {
        fontSize: 13,
        bold: true,
        color: "#FFFFFF",
        alignment: "center",
      });
      addText(slide, `{{RISK_${i + 1}_TITLE}}`, 196, y, 410, 40, {
        fontSize: 20,
        bold: true,
        color: cfg.colors.dark,
      });
      addText(slide, `{{RISK_${i + 1}_IMPACT}}`, 196, y + 42, 980, 58, {
        fontSize: 16,
        color: cfg.colors.muted,
      });
      if (i < 3) addRect(slide, 196, y + 112, 980, 1, cfg.colors.rule, "none");
    }
  }

  // 5. IT recommendations
  {
    const slide = deck.slides.add();
    slide.background.fill = "#FFFFFF";
    await addHeader(slide, cfg, "ИТ-рекомендации укрепляют управляемость среды", 5);
    addBulletRows(slide, ["{{IT_1}}", "{{IT_2}}", "{{IT_3}}", "{{IT_4}}"], cfg, 166, 122);
  }

  // 6. Security recommendations
  {
    const slide = deck.slides.add();
    slide.background.fill = "#FFFFFF";
    await addHeader(slide, cfg, "ИБ-рекомендации снижают наиболее вероятные риски", 6);
    addBulletRows(slide, ["{{SEC_1}}", "{{SEC_2}}", "{{SEC_3}}", "{{SEC_4}}"], cfg, 166, 122);
  }

  // 7. Roadmap
  {
    const slide = deck.slides.add();
    slide.background.fill = "#FFFFFF";
    await addHeader(slide, cfg, "План на 90 дней переводит выводы в действия", 7);
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

  // 8. Decisions
  {
    const slide = deck.slides.add();
    slide.background.fill = cfg.colors.dark;
    addRect(slide, 0, 0, 22, H, cfg.colors.accent, "none");
    await addImage(slide, cfg.logo, { left: 980, top: 36, width: 230, height: 72 }, "contain");
    addText(slide, "Решения, которые стоит принять сейчас", 66, 58, 820, 58, {
      fontSize: 36,
      bold: true,
      color: "#FFFFFF",
    });
    addRect(slide, 66, 134, 112, 8, cfg.colors.accent, "rounded-lg");
    const decisions = ["{{DECISION_1}}", "{{DECISION_2}}", "{{DECISION_3}}", "{{DECISION_4}}"];
    decisions.forEach((item, i) => {
      const y = 184 + i * 92;
      addText(slide, String(i + 1).padStart(2, "0"), 66, y, 48, 30, {
        fontSize: 16,
        bold: true,
        color: cfg.colors.accent,
      });
      addText(slide, item, 130, y - 4, 1020, 68, {
        fontSize: 19,
        color: "#FFFFFF",
      });
    });
    addRect(slide, 66, 582, 570, 74, cfg.colors.accent, "rounded-lg");
    addText(slide, "Следующий шаг: согласовать приоритеты, владельцев и критерии результата", 88, 601, 526, 42, {
      fontSize: 17,
      bold: true,
      color: "#FFFFFF",
      alignment: "center",
    });
    addText(slide, `${cfg.brand} Audit System  |  by Ivan Rudoy`, 760, 617, 450, 24, {
      fontSize: 13,
      bold: true,
      color: "#FFFFFF",
      alignment: "right",
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
      logo: path.join(STATIC, "presentation_khalil_logo.png"),
      cover: path.join(STATIC, "presentation_khalil_cover.png"),
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
      logo: path.join(STATIC, "presentation_btg_logo.png"),
      cover: path.join(STATIC, "presentation_btg_cover.png"),
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
