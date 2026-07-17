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
  const positions = [{ x: 58, y: 145 }, { x: 58, y: 392 }];
  for (let index = 0; index < 2; index += 1) {
    const number = startIndex + index;
    const { x, y } = positions[index];
    addRect(slide, x, y, 1120, 218, "#FFFFFF", "rounded-lg", {
      style: "solid",
      fill: cfg.colors.rule,
      width: 1,
    });
    addRect(slide, x, y, 118, 218, `#B1000${number}`, "rounded-lg");
    addText(slide, `{{REC_${number}_LEVEL}}`, x + 12, y + 26, 94, 24, {
      fontSize: 12,
      bold: true,
      color: `#B2000${number}`,
      alignment: "center",
    });
    addText(slide, String(number).padStart(2, "0"), x + 20, y + 82, 78, 62, {
      fontSize: 34,
      bold: true,
      color: `#B2000${number}`,
      alignment: "center",
    });
    addText(slide, `{{REC_${number}_TITLE}}`, x + 144, y + 15, 512, 34, {
      fontSize: 19,
      bold: true,
      color: cfg.colors.dark,
    });
    addText(slide, "Основание приоритета", x + 144, y + 58, 180, 18, { fontSize: 11, bold: true, color: cfg.colors.accent });
    addText(slide, `{{REC_${number}_EVIDENCE}}`, x + 144, y + 78, 470, 46, {
      fontSize: 13,
      color: cfg.colors.muted,
    });
    addText(slide, "Что сделать", x + 144, y + 132, 120, 18, { fontSize: 11, bold: true, color: cfg.colors.accent });
    addText(slide, `{{REC_${number}_ACTION}}`, x + 144, y + 152, 470, 52, {
      fontSize: 13,
      bold: true,
      color: cfg.colors.dark,
    });
    addRect(slide, x + 636, y + 18, 1, 182, cfg.colors.rule, "none");
    addText(slide, "Решение и производители", x + 660, y + 18, 420, 18, { fontSize: 11, bold: true, color: cfg.colors.accent });
    addText(slide, `{{REC_${number}_SOLUTION}} · {{REC_${number}_VENDORS}}`, x + 660, y + 40, 420, 45, {
      fontSize: 13,
      bold: true,
      color: cfg.colors.dark,
    });
    addText(slide, "Основание", x + 660, y + 95, 120, 18, { fontSize: 11, bold: true, color: cfg.colors.accent });
    addText(slide, `{{REC_${number}_LEGAL}}`, x + 660, y + 116, 420, 36, { fontSize: 12, color: cfg.colors.muted });
    addText(slide, "Критерий результата", x + 660, y + 160, 150, 18, { fontSize: 11, bold: true, color: cfg.colors.accent });
    addText(slide, `{{REC_${number}_METRIC}}`, x + 660, y + 181, 420, 28, { fontSize: 12, bold: true, color: cfg.colors.dark });
  }
}

function addEllipse(slide, x, y, w, h, fill, line = "none") {
  return slide.shapes.add({
    geometry: "ellipse",
    position: { left: x, top: y, width: w, height: h },
    fill,
    line: line === "none" ? { style: "solid", fill: "none", width: 0 } : line,
  });
}

async function buildTemplate(cfg) {
  const deck = Presentation.create({ slideSize: { width: W, height: H } });

  // 1. Cover
  {
    const slide = deck.slides.add();
    slide.background.fill = "#FFFFFF";
    addRect(slide, 0, 0, 24, H, cfg.colors.accent, "none");
    addRect(slide, 24, 0, 736, H, cfg.colors.soft, "none");
    await addImage(slide, cfg.cover, { left: 760, top: 0, width: 520, height: H }, "cover");
    addRect(slide, 760, 0, 10, H, cfg.colors.accent, "none");
    await addImage(slide, cfg.logo, { left: 72, top: 54, width: 250, height: 88 }, "contain");
    addText(slide, "ЭКСПЕРТНЫЙ АУДИТ ИТ И ИБ", 72, 182, 610, 30, {
      fontSize: 16,
      bold: true,
      color: cfg.colors.accent,
    });
    addText(slide, "{{COMPANY}}", 72, 230, 620, 150, {
      fontSize: 48,
      bold: true,
      color: cfg.colors.dark,
    });
    addText(slide, "{{INDUSTRY}}  •  {{CITY}}", 72, 390, 620, 52, {
      fontSize: 20,
      color: cfg.colors.muted,
    });
    addRect(slide, 72, 500, 198, 104, "#C10001", "rounded-xl", { style: "solid", fill: cfg.colors.rule, width: 1 });
    addText(slide, "Зрелость ИТ", 88, 518, 166, 24, {
      fontSize: 14,
      bold: true,
      color: "#C20001",
      alignment: "center",
    });
    addText(slide, "{{IT_SCORE}}%", 88, 548, 166, 42, {
      fontSize: 32,
      bold: true,
      color: "#C20001",
      alignment: "center",
    });
    addRect(slide, 286, 500, 198, 104, "#C10002", "rounded-xl", { style: "solid", fill: cfg.colors.rule, width: 1 });
    addText(slide, "Зрелость ИБ", 302, 518, 166, 24, {
      fontSize: 14,
      bold: true,
      color: "#C20002",
      alignment: "center",
    });
    addText(slide, "{{SCORE}}%", 302, 548, 166, 42, {
      fontSize: 32,
      bold: true,
      color: "#C20002",
      alignment: "center",
    });
    addText(slide, "{{DATE}}", 520, 552, 174, 30, {
      fontSize: 16,
      bold: true,
      color: cfg.colors.muted,
      alignment: "right",
    });
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

    // The company profile is intentionally moved near the close. This slide is the executive snapshot.
    addRect(slide, 0, 0, W, H, "#FFFFFF", "none");
    await addHeader(slide, cfg, "Аудит в одном экране", 2);
    addText(slide, "{{SUMMARY_TITLE}}", 58, 150, 720, 96, { fontSize: 29, bold: true, color: cfg.colors.dark });
    addRect(slide, 840, 148, 158, 126, "#C10001", "rounded-lg");
    addText(slide, "ЗРЕЛОСТЬ ИТ", 858, 170, 122, 20, { fontSize: 12, bold: true, color: "#C20001", alignment: "center" });
    addText(slide, "{{IT_SCORE}}%", 858, 201, 122, 46, { fontSize: 34, bold: true, color: "#C20001", alignment: "center" });
    addRect(slide, 1018, 148, 158, 126, "#C10002", "rounded-lg");
    addText(slide, "ЗРЕЛОСТЬ ИБ", 1036, 170, 122, 20, { fontSize: 12, bold: true, color: "#C20002", alignment: "center" });
    addText(slide, "{{SCORE}}%", 1036, 201, 122, 46, { fontSize: 34, bold: true, color: "#C20002", alignment: "center" });
    addText(slide, "Ключевые наблюдения", 58, 306, 300, 30, { fontSize: 19, bold: true, color: cfg.colors.accent });
    addBulletRows(slide, ["{{SUMMARY_1}}", "{{SUMMARY_2}}", "{{SUMMARY_3}}"], cfg, 356, 92);
  }

  // 3. Target outcomes
  {
    const slide = deck.slides.add();
    slide.background.fill = "#FFFFFF";
    await addHeader(slide, cfg, "Что должно измениться за первые 90 дней", 3);
    addText(slide, "Четыре приоритета связывают факты анкеты с измеримым результатом", 58, 140, 1118, 42, {
      fontSize: 20,
      bold: true,
      color: cfg.colors.accent,
    });
    addText(slide, "ПРИОРИТЕТ", 58, 190, 270, 22, { fontSize: 11, bold: true, color: cfg.colors.muted });
    addText(slide, "СЕЙЧАС", 350, 190, 330, 22, { fontSize: 11, bold: true, color: cfg.colors.muted });
    addText(slide, "ЦЕЛЬ 90 ДНЕЙ", 760, 190, 416, 22, { fontSize: 11, bold: true, color: cfg.colors.muted });
    for (let i = 0; i < 4; i += 1) {
      const y = 226 + i * 108;
      addText(slide, String(i + 1).padStart(2, "0"), 58, y + 4, 44, 28, { fontSize: 15, bold: true, color: cfg.colors.accent });
      addText(slide, `{{OUTCOME_${i + 1}_TITLE}}`, 110, y, 218, 68, { fontSize: 17, bold: true, color: cfg.colors.dark });
      addText(slide, `{{OUTCOME_${i + 1}_FROM}}`, 350, y, 330, 72, { fontSize: 14, color: cfg.colors.muted });
      addText(slide, "→", 700, y + 12, 40, 36, { fontSize: 26, bold: true, color: cfg.colors.accent, alignment: "center" });
      addText(slide, `{{OUTCOME_${i + 1}_TO}}`, 760, y, 416, 72, { fontSize: 15, bold: true, color: cfg.colors.dark });
      if (i < 3) addRect(slide, 58, y + 88, 1118, 1, cfg.colors.rule, "none");
    }
  }

  // 4. Infrastructure profile
  {
    const slide = deck.slides.add();
    slide.background.fill = "#FFFFFF";
    await addHeader(slide, cfg, "Контекст среды и уже работающие контроли", 4);
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
    addText(slide, "Подтвержденные сильные стороны", 58, 432, 360, 28, { fontSize: 18, bold: true, color: cfg.colors.accent });
    for (let i = 0; i < 3; i += 1) {
      const x = 58 + i * 382;
      addRect(slide, x, 482, 7, 140, cfg.colors.accent, "rounded-lg");
      addText(slide, `{{STRENGTH_${i + 1}}}`, x + 24, 490, 326, 112, { fontSize: 17, bold: true, color: cfg.colors.dark });
    }
  }

  // 5. Risks
  {
    const slide = deck.slides.add();
    slide.background.fill = "#FFFFFF";
    await addHeader(slide, cfg, "Карта подтвержденных рисков", 5);
    for (let i = 0; i < 6; i += 1) {
      const col = i % 2;
      const row = Math.floor(i / 2);
      const x = 58 + col * 578;
      const y = 146 + row * 164;
      addRect(slide, x, y, 542, 142, "#FFFFFF", "rounded-lg", { style: "solid", fill: cfg.colors.rule, width: 1 });
      addRect(slide, x, y, 108, 142, `#A1000${i + 1}`, "rounded-lg");
      addText(slide, `{{RISK_${i + 1}_LEVEL}}`, x + 10, y + 54, 88, 28, { fontSize: 12, bold: true, color: `#A2000${i + 1}`, alignment: "center" });
      addText(slide, `{{RISK_${i + 1}_TITLE}}`, x + 128, y + 16, 388, 42, { fontSize: 17, bold: true, color: cfg.colors.dark });
      addText(slide, `{{RISK_${i + 1}_IMPACT}}`, x + 128, y + 66, 388, 60, { fontSize: 13, color: cfg.colors.muted });
    }
  }

  // 6. Sector applicability
  {
    const slide = deck.slides.add();
    slide.background.fill = "#FFFFFF";
    await addHeader(slide, cfg, "Почему требования применимы к вашей организации", 6);
    addText(slide, "{{REG_TITLE}}", 58, 144, 1118, 76, { fontSize: 27, bold: true, color: cfg.colors.accent });
    const columns = [
      ["01", "Основание профиля", "{{REG_APPLICABILITY}}"],
      ["02", "Что требуется подтвердить", "{{REG_EXPECTATIONS}}"],
      ["03", "Как читать рекомендации", "{{REG_IMPLEMENTATION}}"],
    ];
    columns.forEach(([number, title, text], i) => {
      const x = 58 + i * 382;
      if (i > 0) addRect(slide, x - 28, 252, 1, 246, cfg.colors.rule, "none");
      addText(slide, number, x, 254, 46, 28, { fontSize: 15, bold: true, color: cfg.colors.accent });
      addText(slide, title, x, 292, 326, 54, { fontSize: 20, bold: true, color: cfg.colors.dark });
      addText(slide, text, x, 360, 326, 126, { fontSize: 16, color: cfg.colors.muted });
    });
    addRect(slide, 58, 536, 1118, 60, cfg.colors.soft, "rounded-lg", { style: "solid", fill: cfg.colors.rule, width: 1 });
    addText(slide, "Нормы Республики Казахстан", 82, 548, 230, 18, { fontSize: 11, bold: true, color: cfg.colors.accent });
    addText(slide, "{{REG_ANCHORS}}", 82, 569, 1070, 18, { fontSize: 12, bold: true, color: cfg.colors.dark });
    addRect(slide, 58, 606, 1118, 58, "#FFFFFF", "rounded-lg", { style: "solid", fill: cfg.colors.rule, width: 1 });
    addText(slide, "Стандарты при подтверждении применимости", 82, 617, 330, 18, { fontSize: 11, bold: true, color: cfg.colors.accent });
    addText(slide, "{{FRAMEWORKS}}", 82, 638, 1070, 18, { fontSize: 12, bold: true, color: cfg.colors.dark });
    addText(slide, "PCI DSS и GDPR применяются только при наличии соответствующих данных и операций.", 58, 671, 1118, 16, { fontSize: 10, color: cfg.colors.muted });
  }

  // 7-10. Detailed recommendations
  {
    const slide = deck.slides.add();
    slide.background.fill = "#FFFFFF";
    await addHeader(slide, cfg, "Рекомендации: первоочередные меры", 7);
    addRecommendationCards(slide, 1, cfg);
  }

  // 7. Next recommendations
  {
    const slide = deck.slides.add();
    slide.background.fill = "#FFFFFF";
    await addHeader(slide, cfg, "Рекомендации: снижение основных рисков", 8);
    addRecommendationCards(slide, 3, cfg);
  }
  {
    const slide = deck.slides.add();
    slide.background.fill = "#FFFFFF";
    await addHeader(slide, cfg, "Рекомендации: устойчивость ИТ и ИБ", 9);
    addRecommendationCards(slide, 5, cfg);
  }
  {
    const slide = deck.slides.add();
    slide.background.fill = "#FFFFFF";
    await addHeader(slide, cfg, "Рекомендации: следующий уровень зрелости", 10);
    addRecommendationCards(slide, 7, cfg);
  }

  // 11. Roadmap timeline
  {
    const slide = deck.slides.add();
    slide.background.fill = "#FFFFFF";
    await addHeader(slide, cfg, "План на 90 дней переводит выводы в действия", 11);
    const phases = [
      ["0–30 дней", "ПРИОРИТЕТЫ И БЫСТРЫЕ МЕРЫ", "{{ROADMAP_1_1}}", "{{ROADMAP_1_2}}", "{{ROADMAP_1_RESULT}}"],
      ["31–60 дней", "ПИЛОТЫ И РЕГЛАМЕНТЫ", "{{ROADMAP_2_1}}", "{{ROADMAP_2_2}}", "{{ROADMAP_2_RESULT}}"],
      ["61–90 дней", "МАСШТАБИРОВАНИЕ И КОНТРОЛЬ", "{{ROADMAP_3_1}}", "{{ROADMAP_3_2}}", "{{ROADMAP_3_RESULT}}"],
    ];
    addText(slide, "Сначала 20% первоочередных мер, которые закрывают основную долю подтвержденных рисков.", 58, 142, 1118, 34, { fontSize: 18, color: cfg.colors.muted });
    addRect(slide, 150, 226, 920, 5, cfg.colors.rule, "none");
    phases.forEach(([phase, label, one, two, result], i) => {
      const x = 58 + i * 382;
      const nodeX = x + 151;
      addText(slide, phase, x, 184, 354, 28, { fontSize: 17, bold: true, color: i === 0 ? cfg.colors.accent : cfg.colors.dark, alignment: "center" });
      addEllipse(slide, nodeX, 202, 52, 52, i === 0 ? cfg.colors.accent : cfg.colors.dark);
      addText(slide, String(i + 1), nodeX + 9, 213, 34, 26, { fontSize: 16, bold: true, color: "#FFFFFF", alignment: "center" });
      addText(slide, label, x + 10, 276, 334, 46, { fontSize: 13, bold: true, color: cfg.colors.accent, alignment: "center" });
      addText(slide, "01", x + 10, 342, 34, 24, { fontSize: 13, bold: true, color: cfg.colors.accent });
      addText(slide, one, x + 52, 336, 292, 86, { fontSize: 15, color: cfg.colors.dark });
      addText(slide, "02", x + 10, 438, 34, 24, { fontSize: 13, bold: true, color: cfg.colors.accent });
      addText(slide, two, x + 52, 432, 292, 86, { fontSize: 15, color: cfg.colors.dark });
      addRect(slide, x + 10, 542, 334, 80, cfg.colors.soft, "rounded-lg", { style: "solid", fill: cfg.colors.rule, width: 1 });
      addText(slide, "ЧТО ДАСТ ЭТАП", x + 26, 555, 302, 18, { fontSize: 10, bold: true, color: cfg.colors.accent });
      addText(slide, result, x + 26, 579, 302, 34, { fontSize: 13, bold: true, color: cfg.colors.dark });
      if (i < 2) addRect(slide, x + 369, 276, 1, 346, cfg.colors.rule, "none");
    });
  }

  // 12. Company profile near the close
  {
    const slide = deck.slides.add();
    slide.background.fill = "#FFFFFF";
    await addHeader(slide, cfg, "Команда для реализации изменений", 12);
    addText(slide, cfg.companyName, 58, 150, 540, 44, { fontSize: 27, bold: true, color: cfg.colors.dark });
    addText(slide, cfg.companySummary, 58, 212, 540, 82, { fontSize: 18, color: cfg.colors.muted });
    addText(slide, "От оценки и проектирования до внедрения, поддержки и развития.", 58, 324, 540, 70, { fontSize: 22, bold: true, color: cfg.colors.accent });
    const stats = [[cfg.foundedYear, "год основания"], ["70+", "крупных проектов"], ["40+", "технологических партнеров"]];
    stats.forEach(([value, label], i) => {
      const x = 58 + i * 180;
      addText(slide, value, x, 456, 150, 48, { fontSize: 31, bold: true, color: cfg.colors.dark, alignment: "center" });
      addText(slide, label, x, 510, 150, 50, { fontSize: 13, color: cfg.colors.muted, alignment: "center" });
    });
    addRect(slide, 660, 146, 520, 466, cfg.colors.soft, "rounded-lg");
    addText(slide, "Компетенции", 700, 178, 430, 34, { fontSize: 23, bold: true, color: cfg.colors.dark });
    ["Аудит и информационная безопасность", "Сетевая и серверная инфраструктура", "Автоматизация и цифровые решения", "Внедрение, поддержка и передача знаний"].forEach((item, i) => {
      const y = 248 + i * 82;
      addText(slide, String(i + 1).padStart(2, "0"), 700, y, 42, 26, { fontSize: 14, bold: true, color: cfg.colors.accent });
      addText(slide, item, 754, y - 3, 370, 52, { fontSize: 17, bold: true, color: cfg.colors.dark });
    });
  }

  // 13. Decisions and call to action
  {
    const slide = deck.slides.add();
    slide.background.fill = "#FFFFFF";
    await addHeader(slide, cfg, "Зафиксируйте решения и следующий шаг", 13);
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
      cover: path.join(STATIC, "presentation_audit_cover.jpg"),
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
      cover: path.join(STATIC, "presentation_audit_cover.jpg"),
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
