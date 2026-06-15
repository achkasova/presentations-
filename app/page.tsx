"use client";

import {
  Ban,
  BarChart3,
  ChevronDown,
  CircleHelp,
  Clipboard,
  Download,
  FileText,
  History,
  Image as ImageIcon,
  Laptop,
  Library,
  Maximize2,
  Monitor,
  Moon,
  Pencil,
  Plus,
  RotateCcw,
  Sparkles,
  Smartphone,
  Sun,
  Trash2,
  Undo2,
  X,
} from "lucide-react";
import PptxGenJS from "pptxgenjs";
import { ChangeEvent, DragEvent, ReactNode, useMemo, useRef, useState } from "react";

type PresentationType =
  | "Демо проекта"
  | "Kickoff"
  | "Бриф"
  | "Воркшоп"
  | "Статус проекта"
  | "Исследование"
  | "Коммерческое предложение";

type PresentationStyle =
  | "Umbrella корпоративный"
  | "Нейтральный"
  | "Клиентский";

type SlideCount = "5–7" | "8–10" | "11–15";
type SlideTheme = "dark" | "light";
type MockupType = "iphone" | "browser" | "macbook";
type MockupChoice = MockupType | "none";

type TextSource = "text" | "file";

type UploadedFile = {
  name: string;
  size: number;
  type: string;
};

type UploadedImage = UploadedFile & {
  id: string;
  dataUrl: string;
};

type ChartType =
  | "ai"
  | "bar"
  | "line"
  | "pie"
  | "donut"
  | "area"
  | "funnel"
  | "timeline"
  | "radar";

type Slide = {
  title: string;
  kind: "cover" | "agenda" | "content" | "metrics" | "image" | "final";
  body?: string;
  bullets?: string[];
};

type SlideLayout = {
  imageX: number;
  imageY: number;
  imageScale: number;
  imageTransparency: number;
  textX: number;
  textY: number;
};

type DeckSnapshot = {
  slides: Slide[];
  slideThemes: Record<number, SlideTheme>;
  slideImages: Record<number, string>;
  slideLayouts: Record<number, SlideLayout>;
};

type PresentationPrompt = {
  goal: string;
  structure: string[];
  rules: string[];
};

const presentationTypes: PresentationType[] = [
  "Демо проекта",
  "Kickoff",
  "Бриф",
  "Воркшоп",
  "Статус проекта",
  "Исследование",
  "Коммерческое предложение",
];

const styles: PresentationStyle[] = [
  "Umbrella корпоративный",
  "Нейтральный",
  "Клиентский",
];

const slideCounts: SlideCount[] = ["5–7", "8–10", "11–15"];
const slideThemeOptions: Array<{ label: string; value: SlideTheme }> = [
  { label: "Тёмная", value: "dark" },
  { label: "Светлая", value: "light" },
];

const mockupOptions: Array<{
  label: string;
  value: MockupChoice;
  icon: typeof Smartphone;
}> = [
  { label: "Без мокапа", value: "none", icon: Ban },
  { label: "iPhone", value: "iphone", icon: Smartphone },
  { label: "Browser", value: "browser", icon: Monitor },
  { label: "Macbook", value: "macbook", icon: Laptop },
];

const standardImages: UploadedImage[] = Array.from({ length: 14 }, (_, index) => ({
  id: `standard-${index + 1}`,
  name: `Иллюстрация ${index + 1}`,
  size: 0,
  type: "image/png",
  dataUrl: `/standard-images/standard-${String(index + 1).padStart(2, "0")}.png`,
}));

const chartTypeOptions: Array<{ label: string; value: ChartType }> = [
  { label: "Выбор AI", value: "ai" },
  { label: "Столбцы", value: "bar" },
  { label: "Линия", value: "line" },
  { label: "Круговая", value: "pie" },
  { label: "Кольцевая", value: "donut" },
  { label: "Область", value: "area" },
  { label: "Воронка", value: "funnel" },
  { label: "Таймлайн", value: "timeline" },
  { label: "Радар", value: "radar" },
];

const demoSlides: Slide[] = [
  { title: "Титульный слайд", kind: "cover", body: "Демо продукта и ключевых сценариев" },
  {
    title: "Повестка",
    kind: "agenda",
    bullets: ["Контекст и задача", "Ключевые сценарии", "Демо экранов", "Результаты и следующие шаги"],
  },
  {
    title: "Контекст и проблема",
    kind: "content",
    bullets: ["Что мешало пользователям", "Почему текущий процесс требует улучшения", "Какой результат важен бизнесу"],
  },
  {
    title: "Цели проекта",
    kind: "content",
    bullets: ["Сократить ручные операции", "Упростить ключевой пользовательский путь", "Подготовить основу для масштабирования"],
  },
  {
    title: "Решение и ключевые функции",
    kind: "content",
    bullets: ["Основной пользовательский сценарий", "Функции, которые показываем в демо", "Отличия от текущего процесса"],
  },
  {
    title: "Результаты тестирования",
    kind: "metrics",
    bullets: ["85% положительная оценка", "60% использование функции", "4.7 средняя оценка"],
  },
  { title: "Что понравилось пользователям", kind: "image" },
  {
    title: "Следующие шаги",
    kind: "content",
    bullets: ["Собрать обратную связь", "Приоритизировать доработки", "Подготовить план релиза"],
  },
  { title: "Финальный слайд", kind: "final", body: "Обсудим вопросы и обратную связь по демо" },
];

const slidePresets: Record<PresentationType, Slide[]> = {
  "Демо проекта": demoSlides,
  Kickoff: [
    { title: "Титульный слайд", kind: "cover", body: "Старт проекта и выравнивание ожиданий" },
    {
      title: "Цели встречи",
      kind: "agenda",
      bullets: ["Зафиксировать цели", "Согласовать роли", "Подтвердить объем работ", "Определить ближайшие шаги"],
    },
    {
      title: "Контекст проекта",
      kind: "content",
      bullets: ["Бизнес-задача", "Ожидаемый результат", "Ограничения и вводные"],
    },
    {
      title: "Команда и роли",
      kind: "content",
      bullets: ["Кто принимает решения", "Кто отвечает за поставку", "Как устроена коммуникация"],
    },
    {
      title: "Объем работ",
      kind: "content",
      bullets: ["Что входит в первый этап", "Что остается вне рамок", "Какие артефакты готовим"],
    },
    {
      title: "План запуска",
      kind: "content",
      bullets: ["Discovery", "Design & validation", "Development", "Release preparation"],
    },
    {
      title: "Риски и зависимости",
      kind: "metrics",
      bullets: ["Доступы и данные", "Согласования с бизнесом", "Интеграции и внешние команды"],
    },
    {
      title: "Следующие шаги",
      kind: "content",
      bullets: ["Подтвердить рабочую группу", "Передать материалы", "Назначить регулярные встречи"],
    },
    { title: "Финальный слайд", kind: "final", body: "Готовы стартовать и синхронизировать команду" },
  ],
  Бриф: [
    { title: "Титульный слайд", kind: "cover", body: "Бриф проекта и первичная постановка задачи" },
    {
      title: "Что нужно выяснить",
      kind: "agenda",
      bullets: ["Бизнес-контекст", "Целевая аудитория", "Ожидаемый результат", "Ограничения и критерии успеха"],
    },
    {
      title: "Задача клиента",
      kind: "content",
      bullets: ["Какую проблему решаем", "Почему она важна сейчас", "Что будет считаться успешным результатом"],
    },
    {
      title: "Аудитория и сценарии",
      kind: "content",
      bullets: ["Кто основные пользователи", "Какие сценарии критичны", "Где возникают сложности"],
    },
    {
      title: "Требования и ограничения",
      kind: "metrics",
      bullets: ["Сроки и бюджет", "Технические зависимости", "Юридические и бренд-ограничения"],
    },
    {
      title: "Открытые вопросы",
      kind: "content",
      bullets: ["Каких данных не хватает", "Что нужно уточнить у стейкхолдеров", "Какие решения еще не приняты"],
    },
    {
      title: "Следующие шаги",
      kind: "content",
      bullets: ["Собрать ответы", "Подготовить оценку", "Согласовать формат следующей встречи"],
    },
    { title: "Финальный слайд", kind: "final", body: "Зафиксируем вопросы и договоримся о следующем шаге" },
  ],
  Воркшоп: [
    { title: "Титульный слайд", kind: "cover", body: "Рабочая сессия для принятия решений" },
    {
      title: "План воркшопа",
      kind: "agenda",
      bullets: ["Цель сессии", "Карта проблем", "Идеи решений", "Приоритеты и action items"],
    },
    {
      title: "Цель сессии",
      kind: "content",
      bullets: ["Какое решение нужно принять", "Какие вопросы закрываем", "Какие артефакты получим"],
    },
    {
      title: "Участники и роли",
      kind: "content",
      bullets: ["Фасилитатор", "Decision makers", "Эксперты по продукту и технологиям"],
    },
    {
      title: "Проблемное поле",
      kind: "content",
      bullets: ["Главные боли пользователей", "Ограничения процесса", "Гипотезы для проверки"],
    },
    {
      title: "Приоритизация",
      kind: "metrics",
      bullets: ["Impact / Effort", "Быстрые победы", "Риски реализации"],
    },
    {
      title: "Решения и договоренности",
      kind: "content",
      bullets: ["Что выбрали", "Что отложили", "Кто владелец каждого действия"],
    },
    {
      title: "Action items",
      kind: "content",
      bullets: ["Сроки", "Ответственные", "Формат контроля прогресса"],
    },
    { title: "Финальный слайд", kind: "final", body: "Соберем итоги и закрепим договоренности" },
  ],
  "Статус проекта": [
    { title: "Титульный слайд", kind: "cover", body: "Статус проекта и ключевые изменения" },
    {
      title: "Краткий статус",
      kind: "agenda",
      bullets: ["Общий статус", "Что сделано", "Риски и блокеры", "План на следующий период"],
    },
    {
      title: "Прогресс по направлениям",
      kind: "metrics",
      bullets: ["80% задач в работе по плану", "3 блока завершены", "2 риска требуют внимания"],
    },
    {
      title: "Выполнено за период",
      kind: "content",
      bullets: ["Завершенные deliverables", "Принятые решения", "Обновления по командам"],
    },
    {
      title: "Открытые вопросы",
      kind: "content",
      bullets: ["Что требует решения клиента", "Какие данные нужны", "Где нужна синхронизация"],
    },
    {
      title: "Риски",
      kind: "content",
      bullets: ["Сроки", "Зависимости", "Изменение объема работ"],
    },
    {
      title: "План на следующий период",
      kind: "content",
      bullets: ["Приоритеты недели", "Контрольные точки", "Ожидаемые результаты"],
    },
    { title: "Финальный слайд", kind: "final", body: "Зафиксируем решения и следующий отчетный период" },
  ],
  Исследование: [
    { title: "Титульный слайд", kind: "cover", body: "Результаты исследования и продуктовые выводы" },
    {
      title: "Методология",
      kind: "agenda",
      bullets: ["Цель исследования", "Методы и источники", "Выборка", "Как читать выводы"],
    },
    {
      title: "Ключевые наблюдения",
      kind: "content",
      bullets: ["Повторяющиеся паттерны", "Главные боли", "Неочевидные потребности"],
    },
    {
      title: "Сегменты аудитории",
      kind: "metrics",
      bullets: ["3 ключевых сегмента", "2 критичных сценария", "5 гипотез для проверки"],
    },
    {
      title: "Инсайты",
      kind: "content",
      bullets: ["Что влияет на поведение", "Где теряется ценность", "Какие возможности открываются"],
    },
    {
      title: "Рекомендации",
      kind: "content",
      bullets: ["Что изменить в продукте", "Что проверить дополнительно", "Что можно делать уже сейчас"],
    },
    {
      title: "Следующие шаги",
      kind: "content",
      bullets: ["Подтвердить гипотезы", "Обновить roadmap", "Подготовить прототип"],
    },
    { title: "Финальный слайд", kind: "final", body: "Обсудим выводы и приоритеты" },
  ],
  "Коммерческое предложение": [
    { title: "Титульный слайд", kind: "cover", body: "Предложение по решению задачи клиента" },
    {
      title: "Потребность клиента",
      kind: "agenda",
      bullets: ["Контекст", "Цели", "Предлагаемый подход", "План и условия"],
    },
    {
      title: "Предлагаемое решение",
      kind: "content",
      bullets: ["Как решаем задачу", "Почему этот подход подходит", "Какую ценность получает клиент"],
    },
    {
      title: "Объем работ",
      kind: "content",
      bullets: ["Discovery", "Design", "Development", "Support & launch"],
    },
    {
      title: "План и сроки",
      kind: "metrics",
      bullets: ["4 месяца до MVP", "2 недели discovery", "3 контрольные точки"],
    },
    {
      title: "Команда",
      kind: "content",
      bullets: ["Product manager", "UX/UI designer", "Frontend + Backend", "QA"],
    },
    {
      title: "Стоимость и условия",
      kind: "content",
      bullets: ["Формат оценки", "Что входит в стоимость", "Условия старта"],
    },
    { title: "Финальный слайд", kind: "final", body: "Готовы обсудить предложение и следующий шаг" },
  ],
};

const presentationPrompts: Record<PresentationType, PresentationPrompt> = {
  "Демо проекта": {
    goal: "Показать продукт через проблему, решение, ключевые сценарии, экраны и измеримый результат.",
    structure: ["Контекст", "Цели", "Решение", "Демо экранов", "Результаты", "Следующие шаги"],
    rules: [
      "Фокусируйся на том, что зритель увидит в продукте.",
      "Используй скриншоты как доказательства, а не как украшение.",
      "Заканчивай вопросами и сбором обратной связи.",
    ],
  },
  Kickoff: {
    goal: "Выравнять команду проекта по целям, ролям, объему работ, рискам и ближайшему плану.",
    structure: ["Цели встречи", "Контекст", "Команда", "Scope", "План", "Риски", "Next steps"],
    rules: [
      "Не продавай решение, а фиксируй договоренности.",
      "Отделяй in scope от out of scope.",
      "Каждый открытый вопрос превращай в next action.",
    ],
  },
  Бриф: {
    goal: "Структурировать вводные клиента и превратить разрозненный контент в список вопросов, требований и следующих шагов.",
    structure: ["Задача", "Аудитория", "Сценарии", "Ограничения", "Открытые вопросы", "Next steps"],
    rules: [
      "Не додумывай требования, помечай неизвестное как вопрос.",
      "Выделяй критерии успеха и ограничения.",
      "Делай акцент на том, что нужно уточнить до оценки.",
    ],
  },
  Воркшоп: {
    goal: "Подготовить рабочую сессию, где участники принимают решения и получают список action items.",
    structure: ["Цель сессии", "Участники", "Проблемное поле", "Идеи", "Приоритеты", "Договоренности"],
    rules: [
      "Формулируй слайды как рабочие блоки для обсуждения.",
      "Добавляй места для решений, владельцев и сроков.",
      "Не перегружай теорией, выводи группу к action items.",
    ],
  },
  "Статус проекта": {
    goal: "Дать прозрачный отчет о прогрессе, рисках, блокерах и плане на следующий период.",
    structure: ["Executive summary", "Progress", "Done", "Open questions", "Risks", "Plan"],
    rules: [
      "Начинай с общего статуса и изменений с прошлого отчета.",
      "Разделяй факты, риски и решения, которые нужны от клиента.",
      "Каждый риск связывай с влиянием и планом снижения.",
    ],
  },
  Исследование: {
    goal: "Превратить результаты исследования в понятные инсайты, продуктовые выводы и рекомендации.",
    structure: ["Методология", "Наблюдения", "Сегменты", "Инсайты", "Рекомендации", "Next steps"],
    rules: [
      "Отделяй наблюдение от интерпретации.",
      "Подкрепляй инсайты повторяющимися паттернами.",
      "Рекомендации должны вести к продуктовым решениям.",
    ],
  },
  "Коммерческое предложение": {
    goal: "Показать клиенту ценность решения, объем работ, план, команду и условия старта.",
    structure: ["Потребность", "Решение", "Scope", "План", "Команда", "Стоимость", "Next steps"],
    rules: [
      "Связывай каждую часть решения с потребностью клиента.",
      "Пиши уверенно, но без неподтвержденных обещаний.",
      "В финале подводи к следующему коммерческому шагу.",
    ],
  },
};

const defaultBulletsByKind: Record<Slide["kind"], string[]> = {
  cover: ["Building digital products that drive results"],
  agenda: ["Project Overview", "Goals & Success Metrics", "Scope of Work"],
  content: [
    "Упрощает работу команды",
    "Сокращает ручные операции",
    "Помогает быстрее принимать решения",
  ],
  metrics: ["85% положительная оценка", "60% использование функции", "4.7 средняя оценка"],
  image: ["Скриншоты продукта", "Ключевые пользовательские сценарии"],
  final: ["Готовы ответить на ваши вопросы"],
};

const regenerationVariants: Record<Slide["kind"], Array<Pick<Slide, "title" | "body" | "bullets">>> = {
  cover: [
    {
      title: "Project Kickoff",
      body: "Building digital products that drive results",
      bullets: ["Клиентский портал", "Демо-презентация проекта"],
    },
    {
      title: "Building Digital Products That Matter",
      body: "A concise overview of goals, scope and next steps",
      bullets: ["Product demo", "Umbrella IT"],
    },
  ],
  agenda: [
    {
      title: "План презентации",
      bullets: ["Что мы сделали", "Ключевые фичи", "Улучшения и исправления", "Что дальше"],
    },
    {
      title: "Agenda",
      bullets: ["Project Overview", "Goals & Success Metrics", "Timeline", "Next Steps"],
    },
  ],
  content: [
    {
      title: "Что мы сделали",
      body: "Обзор выполненного объема работ и ключевых достижений.",
      bullets: ["Собрали основу продукта", "Подготовили пользовательские сценарии", "Снизили ручные операции"],
    },
    {
      title: "Ключевые результаты",
      body: "Мы сфокусировались на ценности и решении реальных пользовательских задач.",
      bullets: ["2 крупные функции выпущено", "12 улучшений доставлено", "0 критических ошибок"],
    },
    {
      title: "Следующие шаги",
      body: "План работ и ключевые релизы на ближайший период.",
      bullets: ["Review & Sign Off", "Access & Onboarding", "Kickoff Workshop"],
    },
  ],
  metrics: [
    {
      title: "Ключевые метрики",
      bullets: ["98% покрытие тестами", "-40% снижение времени отклика", "+60% рост вовлеченности"],
    },
    {
      title: "Goals & Success Metrics",
      bullets: ["30% increase in user engagement", "20% reduction in manual processes", "4 months to MVP"],
    },
  ],
  image: [
    {
      title: "Материалы и скриншоты",
      body: "Визуальные материалы, которые помогают раскрыть продукт.",
      bullets: ["Ключевые экраны", "Пользовательские сценарии"],
    },
    {
      title: "Feature Demo",
      body: "A walkthrough of the new functionality and user flows.",
      bullets: ["Real-time data updates", "Customizable widgets", "Improved visualization"],
    },
  ],
  final: [
    {
      title: "Спасибо!",
      body: "Готовы ответить на ваши вопросы",
      bullets: ["Questions?", "We're here to help."],
    },
    {
      title: "Q&A",
      body: "Let's discuss your feedback and answer questions.",
      bullets: ["Feedback", "Next steps"],
    },
  ],
};

const hydrateSlide = (slide: Slide): Slide => ({
  ...slide,
  body:
    slide.body ??
    (slide.kind === "cover"
      ? "Building digital products that drive results"
      : slide.kind === "final"
        ? "Готовы ответить на ваши вопросы"
        : "Краткое содержание на основе загруженных материалов."),
  bullets: slide.bullets ?? defaultBulletsByKind[slide.kind],
});

const getSlideLimit = (count: SlideCount) => {
  if (count === "5–7") return 7;
  if (count === "11–15") return 12;
  return 9;
};

const buildPresentationPrompt = ({
  type,
  style,
  count,
  source,
  text,
  fileName,
  hasImages,
}: {
  type: PresentationType;
  style: PresentationStyle;
  count: SlideCount;
  source: TextSource;
  text: string;
  fileName?: string;
  hasImages: boolean;
}) => {
  const scenario = presentationPrompts[type];
  const sourceDescription =
    source === "text" && text.trim()
      ? text.trim().slice(0, 900)
      : fileName
        ? `Контент взят из файла: ${fileName}`
        : "Пользователь не добавил подробный текст. Собери базовую структуру по типу презентации.";

  return [
    `Ты AI-редактор презентаций Umbrella IT. Тип презентации: ${type}.`,
    `Цель: ${scenario.goal}`,
    `Стиль: ${style}. Количество слайдов: ${count}.`,
    `Рекомендуемая структура: ${scenario.structure.join(" -> ")}.`,
    `Правила: ${scenario.rules.join(" ")}`,
    hasImages
      ? "Есть изображения: подбери для них слайды с демонстрацией и не дублируй текстом то, что видно на экране."
      : "Изображений нет: делай структуру самодостаточной без ссылок на скриншоты.",
    `Исходные материалы: ${sourceDescription}`,
    "Верни JSON-массив слайдов: title, kind, body, bullets. Не добавляй маркетинговую воду.",
  ].join("\n");
};

const buildPromptedSlides = (
  type: PresentationType,
  count: SlideCount,
  hasImages: boolean,
) => {
  const limit = getSlideLimit(count);
  const sourceSlides = slidePresets[type].filter(
    (slide) => hasImages || slide.kind !== "image",
  );

  if (sourceSlides.length <= limit) {
    return sourceSlides.map(hydrateSlide);
  }

  const finalSlide = sourceSlides.find((slide) => slide.kind === "final");
  const trimmedSlides = sourceSlides
    .filter((slide) => slide.kind !== "final")
    .slice(0, finalSlide ? limit - 1 : limit);

  return [...trimmedSlides, ...(finalSlide ? [finalSlide] : [])].map(hydrateSlide);
};

const regenerateSlideContent = (slide: Slide, index: number): Slide => {
  const variants = regenerationVariants[slide.kind];
  const variant = variants[index % variants.length];

  return hydrateSlide({
    ...slide,
    ...variant,
  });
};

const formatSize = (bytes: number) => {
  if (bytes < 1024 * 1024) {
    return `${Math.max(1, Math.round(bytes / 1024))} КБ`;
  }

  return `${(bytes / 1024 / 1024).toFixed(1)} МБ`;
};

const acceptedTextFormats = [".docx", ".txt", ".pdf", ".pptx"];
const acceptedImageExtensions = [".png", ".jpg", ".jpeg", ".webp"];
const acceptedImageFormats = ["image/png", "image/jpeg", "image/webp"];
const sourceFileAccept = [...acceptedTextFormats, ...acceptedImageExtensions].join(",");

const isTextMaterialFile = (file: File) => {
  const lowerName = file.name.toLowerCase();
  return acceptedTextFormats.some((format) => lowerName.endsWith(format));
};

const isImageFile = (file: File) => {
  const lowerName = file.name.toLowerCase();
  return (
    acceptedImageFormats.includes(file.type) ||
    acceptedImageExtensions.some((format) => lowerName.endsWith(format))
  );
};

const inferImageMockup = (image: UploadedImage): MockupType => {
  const name = image.name.toLowerCase();

  if (/iphone|phone|mobile|screen|screenshot|моб|телефон/.test(name)) {
    return "iphone";
  }

  if (/dashboard|dash|analytics|chart|report|дашборд|отчет|аналит/.test(name)) {
    return "macbook";
  }

  if (/admin|panel|web|browser|crm|cms|админ|панел/.test(name)) {
    return "browser";
  }

  return "browser";
};

const defaultSlideLayout: SlideLayout = {
  imageX: 0,
  imageY: 0,
  imageScale: 100,
  imageTransparency: 0,
  textX: 0,
  textY: 0,
};

const getMockupOption = (type: MockupChoice) =>
  mockupOptions.find((option) => option.value === type) ?? mockupOptions[1];

const getSlideLayout = (
  layouts: Record<number, SlideLayout>,
  index: number,
): SlideLayout => ({ ...defaultSlideLayout, ...(layouts[index] ?? {}) });

const brandSlideAssets = {
  coverPhoto: "/brand/slides/code-photo-4.jpg",
  agendaPhoto: "/brand/slides/code-photo-2.jpg",
  contentPhoto: "/brand/slides/code-photo-5.jpg",
  metricsIllustration: "/brand/slides/illus-open-source.png",
  rocketIllustration: "/brand/slides/illus-rocket.png",
  questionIllustration: "/brand/slides/illus-question.png",
  discussionIllustration: "/brand/slides/illus-discussion.png",
};

const loadImageDataUrl = async (src: string) => {
  const response = await fetch(src);
  const blob = await response.blob();

  return new Promise<string>((resolve) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result));
    reader.readAsDataURL(blob);
  });
};

const svgToDataUrl = (svg: string) =>
  `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`;

const pickChartType = (description: string, selectedType: ChartType): Exclude<ChartType, "ai"> => {
  if (selectedType !== "ai") return selectedType;

  const text = description.toLowerCase();
  if (/этап|план|timeline|roadmap|срок|квартал/.test(text)) return "timeline";
  if (/ворон|funnel|конвер/.test(text)) return "funnel";
  if (/доля|процент|share|част/.test(text)) return "donut";
  if (/динами|рост|trend|месяц|год/.test(text)) return "line";
  if (/сравн|compare|категор/.test(text)) return "bar";
  return "bar";
};

const createChartImage = (description: string, selectedType: ChartType): UploadedImage => {
  const chartType = pickChartType(description, selectedType);
  const safeTitle =
    description.trim().split(/\n|\.|;/)[0]?.slice(0, 64) || "Диаграмма";
  const title = safeTitle.replace(/[<&>]/g, "");
  const chartShapes: Record<Exclude<ChartType, "ai">, string> = {
    bar: `
      <rect x="86" y="220" width="48" height="120" rx="8" fill="#0050FF"/>
      <rect x="158" y="172" width="48" height="168" rx="8" fill="#88A8FF"/>
      <rect x="230" y="118" width="48" height="222" rx="8" fill="#0050FF"/>
      <rect x="302" y="202" width="48" height="138" rx="8" fill="#C7D7FF"/>
      <path d="M70 340H382" stroke="#D8E2F7" stroke-width="4"/>`,
    line: `
      <path d="M78 302C130 248 156 270 204 206C246 150 288 174 358 104" fill="none" stroke="#0050FF" stroke-width="12" stroke-linecap="round"/>
      <path d="M78 342H382" stroke="#D8E2F7" stroke-width="4"/>
      <circle cx="204" cy="206" r="12" fill="#88A8FF"/><circle cx="358" cy="104" r="12" fill="#0050FF"/>`,
    pie: `
      <path d="M240 112A112 112 0 1 1 146 285L240 224Z" fill="#0050FF"/>
      <path d="M240 112A112 112 0 0 1 352 224H240Z" fill="#88A8FF"/>
      <path d="M352 224A112 112 0 0 1 146 285L240 224Z" fill="#C7D7FF"/>`,
    donut: `
      <circle cx="240" cy="226" r="112" fill="none" stroke="#0050FF" stroke-width="42" stroke-dasharray="430 704" transform="rotate(-90 240 226)"/>
      <circle cx="240" cy="226" r="112" fill="none" stroke="#88A8FF" stroke-width="42" stroke-dasharray="170 704" stroke-dashoffset="-430" transform="rotate(-90 240 226)"/>
      <circle cx="240" cy="226" r="112" fill="none" stroke="#D8E2F7" stroke-width="42" stroke-dasharray="104 704" stroke-dashoffset="-600" transform="rotate(-90 240 226)"/>`,
    area: `
      <path d="M70 322C126 238 168 270 220 186C274 98 318 158 390 86V342H70Z" fill="#0050FF" opacity=".28"/>
      <path d="M70 322C126 238 168 270 220 186C274 98 318 158 390 86" fill="none" stroke="#0050FF" stroke-width="10" stroke-linecap="round"/>`,
    funnel: `
      <path d="M90 110H390L342 176H138Z" fill="#0050FF"/>
      <path d="M142 200H338L302 264H178Z" fill="#88A8FF"/>
      <path d="M184 288H296L268 348H212Z" fill="#C7D7FF"/>`,
    timeline: `
      <path d="M74 226H400" stroke="#0050FF" stroke-width="10" stroke-linecap="round"/>
      <circle cx="102" cy="226" r="24" fill="#0050FF"/><circle cx="224" cy="226" r="24" fill="#88A8FF"/><circle cx="346" cy="226" r="24" fill="#0050FF"/>
      <rect x="78" y="276" width="78" height="34" rx="8" fill="#D8E2F7"/><rect x="198" y="154" width="78" height="34" rx="8" fill="#D8E2F7"/><rect x="320" y="276" width="78" height="34" rx="8" fill="#D8E2F7"/>`,
    radar: `
      <path d="M240 100L354 184L310 326H170L126 184Z" fill="none" stroke="#D8E2F7" stroke-width="6"/>
      <path d="M240 132L320 196L292 292H184L152 198Z" fill="#0050FF" opacity=".35" stroke="#0050FF" stroke-width="8"/>`,
  };
  const svg = `
    <svg xmlns="http://www.w3.org/2000/svg" width="480" height="360" viewBox="0 0 480 360">
      <rect width="480" height="360" rx="28" fill="#F7F9FF"/>
      <rect x="32" y="32" width="416" height="296" rx="22" fill="#FFFFFF"/>
      <rect x="56" y="58" width="68" height="8" rx="4" fill="#0050FF"/>
      <text x="56" y="94" font-family="Inter, Arial" font-size="22" font-weight="800" fill="#282A32">${title}</text>
      ${chartShapes[chartType]}
    </svg>`;

  return {
    id: `chart-${crypto.randomUUID()}`,
    name: title,
    size: svg.length,
    type: "image/svg+xml",
    dataUrl: svgToDataUrl(svg),
  };
};

const buildInitialSlideImages = (slides: Slide[], images: UploadedImage[]) => {
  if (images.length === 0) return {};

  let nextImageIndex = 0;

  return Object.fromEntries(
    slides
      .map((slide, slideIndex) => {
        if (slide.kind !== "image") return null;
        const image = images[nextImageIndex] ?? images[0];
        nextImageIndex += 1;
        return [slideIndex, image.id] as const;
      })
      .filter(Boolean) as Array<readonly [number, string]>,
  );
};

export default function Home() {
  const [textSource, setTextSource] = useState<TextSource>("file");
  const [manualText, setManualText] = useState("");
  const [textFile, setTextFile] = useState<UploadedFile | null>(null);
  const [images, setImages] = useState<UploadedImage[]>([]);
  const [chartImages, setChartImages] = useState<UploadedImage[]>([]);
  const [isStandardLibraryOpen, setIsStandardLibraryOpen] = useState(false);
  const [isChartLibraryOpen, setIsChartLibraryOpen] = useState(false);
  const [chartDescription, setChartDescription] = useState("");
  const [chartType, setChartType] = useState<ChartType>("ai");
  const [presentationType, setPresentationType] =
    useState<PresentationType>("Демо проекта");
  const [style, setStyle] = useState<PresentationStyle>("Umbrella корпоративный");
  const [slideCount, setSlideCount] = useState<SlideCount>("8–10");
  const [deckTheme, setDeckTheme] = useState<SlideTheme>("dark");
  const [slideThemes, setSlideThemes] = useState<Record<number, SlideTheme>>({});
  const [slideImages, setSlideImages] = useState<Record<number, string>>({});
  const [slideLayouts, setSlideLayouts] = useState<Record<number, SlideLayout>>({});
  const [imageMockups, setImageMockups] = useState<Record<string, MockupChoice>>({});
  const [slides, setSlides] = useState<Slide[]>(demoSlides.map(hydrateSlide));
  const [undoStack, setUndoStack] = useState<DeckSnapshot[]>([]);
  const [slideVersions, setSlideVersions] = useState<Record<number, Slide[]>>({});
  const [generationPrompt, setGenerationPrompt] = useState("");
  const [isGenerating, setIsGenerating] = useState(false);
  const [generationError, setGenerationError] = useState("");
  const [hasGenerated, setHasGenerated] = useState(false);
  const [editingSlideIndex, setEditingSlideIndex] = useState<number | null>(null);
  const textFileRef = useRef<HTMLInputElement | null>(null);
  const imageFileRef = useRef<HTMLInputElement | null>(null);
  const chartFileRef = useRef<HTMLInputElement | null>(null);
  const chartImageRef = useRef<HTMLInputElement | null>(null);

  const libraryImages = useMemo(
    () => [...images, ...chartImages, ...standardImages],
    [chartImages, images],
  );

  const materializeSlides = (baseSlides: Slide[]) => {
    if (
      images.length + chartImages.length > 0 &&
      !baseSlides.some((slide) => slide.kind === "image")
    ) {
      const finalSlideIndex = baseSlides.findIndex((slide) => slide.kind === "final");
      const imageSlide: Slide = {
        title: "Скриншот решения",
        kind: "image",
      };

      if (finalSlideIndex === -1) {
        return [...baseSlides, hydrateSlide(imageSlide)];
      }

      return [
        ...baseSlides.slice(0, finalSlideIndex).map(hydrateSlide),
        hydrateSlide(imageSlide),
        ...baseSlides.slice(finalSlideIndex).map(hydrateSlide),
      ];
    }

    return baseSlides.map(hydrateSlide);
  };

  const currentSlides = useMemo(
    () => materializeSlides(slides),
    [chartImages.length, images.length, slides],
  );

  const saveUndoSnapshot = () => {
    setUndoStack((current) =>
      [
        ...current,
        {
          slides: currentSlides,
          slideThemes,
          slideImages,
          slideLayouts,
        },
      ].slice(-12),
    );
  };

  const rememberSlideVersion = (index: number, slide: Slide) => {
    setSlideVersions((current) => ({
      ...current,
      [index]: [slide, ...(current[index] ?? [])].slice(0, 8),
    }));
  };

  const revertLastStep = () => {
    const snapshot = undoStack[undoStack.length - 1];
    if (!snapshot) return;

    setSlides(snapshot.slides.map(hydrateSlide));
    setSlideThemes(snapshot.slideThemes);
    setSlideImages(snapshot.slideImages);
    setSlideLayouts(snapshot.slideLayouts);
    setUndoStack((current) => current.slice(0, -1));
    setHasGenerated(true);
  };

  const handleTextFile = (file?: File) => {
    if (!file) return;

    if (isImageFile(file)) {
      void handleImageFiles([file]);
      return;
    }

    if (!isTextMaterialFile(file)) return;
    setTextFile({ name: file.name, size: file.size, type: file.type });
  };

  const handleImageFiles = async (fileList?: FileList | File[]) => {
    if (!fileList) return;

    const incoming = Array.from(fileList)
      .filter(isImageFile)
      .slice(0, Math.max(0, 10 - images.length));

    const loaded = await Promise.all(
      incoming.map(
        (file) =>
          new Promise<UploadedImage>((resolve) => {
            const reader = new FileReader();
            reader.onload = () =>
              resolve({
                id: `${file.name}-${file.size}-${crypto.randomUUID()}`,
                name: file.name,
                size: file.size,
                type: file.type,
                dataUrl: String(reader.result),
              });
            reader.readAsDataURL(file);
          }),
      ),
    );

    setImages((current) => [...current, ...loaded].slice(0, 10));
  };

  const loadImagesFromFiles = async (fileList?: FileList | File[]) => {
    if (!fileList) return [];

    const incoming = Array.from(fileList).filter(isImageFile);
    return Promise.all(
      incoming.map(
        (file) =>
          new Promise<UploadedImage>((resolve) => {
            const reader = new FileReader();
            reader.onload = () =>
              resolve({
                id: `${file.name}-${file.size}-${crypto.randomUUID()}`,
                name: file.name,
                size: file.size,
                type: file.type,
                dataUrl: String(reader.result),
              });
            reader.readAsDataURL(file);
          }),
      ),
    );
  };

  const handleChartImageFiles = async (fileList?: FileList | File[]) => {
    const loaded = await loadImagesFromFiles(fileList);
    if (loaded.length > 0) {
      setChartImages((current) => [...loaded, ...current].slice(0, 24));
    }
  };

  const handleChartTableFile = (file?: File) => {
    if (!file) return;
    setChartDescription((current) =>
      `${current.trim() ? `${current.trim()}\n` : ""}Источник данных: ${file.name}`,
    );
  };

  const onDropText = (event: DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    handleTextFile(event.dataTransfer.files[0]);
  };

  const onDropImages = (event: DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    void handleImageFiles(event.dataTransfer.files);
  };

  const generateChart = () => {
    const nextChart = createChartImage(chartDescription, chartType);
    setChartImages((current) => [nextChart, ...current].slice(0, 24));
    setChartDescription("");
  };

  const createPresentation = async () => {
    saveUndoSnapshot();
    const nextPrompt = buildPresentationPrompt({
      type: presentationType,
      style,
      count: slideCount,
      source: textSource,
      text: manualText,
      fileName: textFile?.name,
      hasImages: images.length > 0,
    });
    const nextSlides = buildPromptedSlides(
      presentationType,
      slideCount,
      images.length > 0,
    );

    setIsGenerating(true);
    setGenerationError("");
    setGenerationPrompt(nextPrompt);

    try {
      const response = await fetch("/api/generate", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          prompt: nextPrompt,
          fallbackSlides: nextSlides,
        }),
      });

      if (!response.ok) {
        throw new Error("AI generation failed");
      }

      const data = (await response.json()) as {
        slides?: Slide[];
        warning?: string;
      };
      const generatedSlides =
        data.slides && data.slides.length > 0
          ? data.slides.map(hydrateSlide)
          : nextSlides;

      setSlides(generatedSlides);
      setSlideThemes(
        Object.fromEntries(generatedSlides.map((_, index) => [index, deckTheme])),
      );
      setSlideImages(buildInitialSlideImages(generatedSlides, images));
      setSlideLayouts({});
      setHasGenerated(true);

      if (data.warning) {
        setGenerationError("AI вернул нестандартный ответ, использован локальный шаблон.");
      }
    } catch {
      setSlides(nextSlides);
      setSlideThemes(
        Object.fromEntries(nextSlides.map((_, index) => [index, deckTheme])),
      );
      setSlideImages(buildInitialSlideImages(nextSlides, images));
      setSlideLayouts({});
      setHasGenerated(true);
      setGenerationError(
        "AI недоступен: проверьте OPENAI_API_KEY. Пока использован локальный шаблон.",
      );
    } finally {
      setIsGenerating(false);
    }
  };

  const getSlideTheme = (index: number): SlideTheme =>
    slideThemes[index] ?? deckTheme;

  const updateSlideTheme = (index: number, theme: SlideTheme) => {
    saveUndoSnapshot();
    setSlideThemes((current) => ({ ...current, [index]: theme }));
  };

  const toggleSlideTheme = (index: number) => {
    updateSlideTheme(index, getSlideTheme(index) === "dark" ? "light" : "dark");
  };

  const getImageMockup = (image: UploadedImage): MockupChoice =>
    imageMockups[image.id] ?? "none";

  const getImageById = (imageId?: string) =>
    libraryImages.find((image) => image.id === imageId);

  const getSlideImage = (index: number) => {
    const assignedImage = getImageById(slideImages[index]);
    if (assignedImage) return assignedImage;

    return currentSlides[index]?.kind === "image"
      ? images[0] ?? chartImages[0]
      : undefined;
  };

  const getSlideImageMockup = (index: number) => {
    const image = getSlideImage(index);
    return image ? getImageMockup(image) : undefined;
  };

  const updateImageMockup = (imageId: string, mockup: MockupChoice) => {
    setImageMockups((current) => ({ ...current, [imageId]: mockup }));
  };

  const removeImage = (imageId: string) => {
    setImages((current) => current.filter((item) => item.id !== imageId));
    setChartImages((current) => current.filter((item) => item.id !== imageId));
    setSlideImages((current) =>
      Object.fromEntries(
        Object.entries(current).filter(([, assignedImageId]) => assignedImageId !== imageId),
      ),
    );
    setImageMockups((current) => {
      const next = { ...current };
      delete next[imageId];
      return next;
    });
  };

  const replaceSlideAt = (index: number, nextSlide: Slide) => {
    const previousSlide = currentSlides[index];
    saveUndoSnapshot();
    if (previousSlide) {
      rememberSlideVersion(index, previousSlide);
    }

    setSlides((current) => {
      const materialized = materializeSlides(current);
      return materialized.map((slide, slideIndex) =>
        slideIndex === index ? hydrateSlide(nextSlide) : slide,
      );
    });
  };

  const updateSlideText = (
    index: number,
    patch: Partial<Pick<Slide, "title" | "body" | "bullets">>,
  ) => {
    const currentSlide = currentSlides[index];
    if (!currentSlide) return;
    replaceSlideAt(index, { ...currentSlide, ...patch });
  };

  const regenerateSlide = (index: number) => {
    const currentSlide = currentSlides[index];
    if (!currentSlide) return;
    replaceSlideAt(index, regenerateSlideContent(currentSlide, index + Date.now()));
  };

  const addSlideAfter = (index: number) => {
    saveUndoSnapshot();
    const newSlide: Slide = hydrateSlide({
      title: "Новый слайд",
      kind: "content",
      bullets: ["Ключевая мысль", "Подтверждение", "Следующий шаг"],
    });
    setSlides((current) => {
      const materialized = materializeSlides(current);
      return [
        ...materialized.slice(0, index + 1),
        newSlide,
        ...materialized.slice(index + 1),
      ];
    });
    setHasGenerated(true);
  };

  const deleteSlideAt = (index: number) => {
    if (currentSlides.length <= 1) return;
    const currentSlide = currentSlides[index];
    saveUndoSnapshot();
    if (currentSlide) rememberSlideVersion(index, currentSlide);
    setSlides((current) => {
      const materialized = materializeSlides(current);
      return materialized.filter((_, slideIndex) => slideIndex !== index);
    });
    setSlideImages((current) =>
      Object.fromEntries(
        Object.entries(current)
          .map(([key, value]) => [Number(key), value] as const)
          .filter(([key]) => key !== index)
          .map(([key, value]) => [key > index ? key - 1 : key, value]),
      ),
    );
    setSlideLayouts((current) =>
      Object.fromEntries(
        Object.entries(current)
          .map(([key, value]) => [Number(key), value] as const)
          .filter(([key]) => key !== index)
          .map(([key, value]) => [key > index ? key - 1 : key, value]),
      ),
    );
  };

  const assignImageToSlide = (index: number, imageId: string) => {
    const currentSlide = currentSlides[index];
    if (!currentSlide || !getImageById(imageId)) return;

    saveUndoSnapshot();
    setSlideImages((current) => ({ ...current, [index]: imageId }));
    if (currentSlide.kind !== "image") {
      replaceSlideAt(index, { ...currentSlide, kind: "image" });
    }
  };

  const pasteImageToSlide = async (index: number) => {
    if (!navigator.clipboard?.read) return;

    try {
      const clipboardItems = await navigator.clipboard.read();
      for (const item of clipboardItems) {
        const imageType = item.types.find((type) => type.startsWith("image/"));
        if (!imageType) continue;

        const blob = await item.getType(imageType);
        const image = await new Promise<UploadedImage>((resolve) => {
          const reader = new FileReader();
          reader.onload = () =>
            resolve({
              id: `clipboard-${crypto.randomUUID()}`,
              name: "Из буфера обмена",
              size: blob.size,
              type: imageType,
              dataUrl: String(reader.result),
            });
          reader.readAsDataURL(blob);
        });
        setImages((current) => [image, ...current].slice(0, 10));
        saveUndoSnapshot();
        setSlideImages((current) => ({ ...current, [index]: image.id }));
        const currentSlide = currentSlides[index];
        if (currentSlide && currentSlide.kind !== "image") {
          replaceSlideAt(index, { ...currentSlide, kind: "image" });
        }
        return;
      }
    } catch {
      setGenerationError("Не удалось прочитать буфер обмена. Разрешите доступ браузеру.");
    }
  };

  const updateSlideLayout = (index: number, patch: Partial<SlideLayout>) => {
    setSlideLayouts((current) => ({
      ...current,
      [index]: {
        ...getSlideLayout(current, index),
        ...patch,
      },
    }));
  };

  const commitSlideLayoutChange = () => {
    saveUndoSnapshot();
  };

  const restoreSlideVersion = (index: number, version: Slide) => {
    replaceSlideAt(index, version);
  };

  const downloadPptx = async () => {
    const pptx = new PptxGenJS();
    pptx.layout = "LAYOUT_WIDE";
    pptx.author = "Umbrella Deck Builder";
    pptx.subject = `${presentationType} — ${style}`;
    pptx.title = "Umbrella Presentation";
    pptx.company = "Umbrella";
    pptx.theme = {
      headFontFace: "ALS Sector",
      bodyFontFace: "Graphik LC",
    };

    const darkPalette = {
      dark: "07111C",
      dark2: "0B1724",
      panel: "101D2B",
      panel2: "142234",
      blue: "0050FF",
      blue2: "0E6BFF",
      soft: "CFDEFF",
      white: "FFFFFF",
      muted: "B8C3D4",
      line: "263545",
    };
    const lightPalette = {
      dark: "F7F7F7",
      dark2: "FFFFFF",
      panel: "FFFFFF",
      panel2: "F2F6FF",
      blue: "0050FF",
      blue2: "0E6BFF",
      soft: "0050FF",
      white: "282A32",
      muted: "6B6D75",
      line: "B4B5B7",
    };
    let C = darkPalette;

    const addBase = (page: PptxGenJS.Slide, pageNumber: number) => {
      page.background = { color: C.dark };
      page.addShape(pptx.ShapeType.rect, {
        x: 0,
        y: 0,
        w: 13.33,
        h: 7.5,
        line: { color: C.dark, transparency: 100 },
        fill: { color: C.dark },
      });
      page.addShape(pptx.ShapeType.rect, {
        x: 0,
        y: 0,
        w: 13.33,
        h: 7.5,
        line: { color: C.dark, transparency: 100 },
        fill: { color: C.blue, transparency: 94 },
      });
      page.addShape(pptx.ShapeType.line, {
        x: 9.2,
        y: 0.65,
        w: 3.4,
        h: 1.1,
        line: { color: C.blue, transparency: 72, width: 1.2 },
      });
      page.addShape(pptx.ShapeType.line, {
        x: 9.85,
        y: 5.8,
        w: 2.7,
        h: -1.25,
        line: { color: C.blue, transparency: 82, width: 1 },
      });
      page.addText("Umbrella IT", {
        x: 0.55,
        y: 0.36,
        w: 1.35,
        h: 0.24,
        color: C.white,
        fontFace: "Graphik LC",
        bold: true,
        fontSize: 9,
        margin: 0,
      });
      page.addShape(pptx.ShapeType.arc, {
        x: 0.36,
        y: 0.37,
        w: 0.18,
        h: 0.18,
        line: { color: C.blue, width: 1 },
        fill: { color: C.blue },
      });
      page.addShape(pptx.ShapeType.line, {
        x: 0.36,
        y: 6.94,
        w: 0.34,
        h: 0,
        line: { color: C.line, width: 0.5 },
      });
      page.addText(String(pageNumber), {
        x: 0.38,
        y: 6.86,
        w: 0.28,
        h: 0.18,
        color: C.muted,
        fontSize: 6,
        margin: 0,
      });
    };

    const addBlueRule = (page: PptxGenJS.Slide, x = 0.68, y = 0.78) => {
      page.addShape(pptx.ShapeType.rect, {
        x,
        y,
        w: 0.7,
        h: 0.04,
        line: { color: C.blue },
        fill: { color: C.blue },
      });
    };

    const addTitle = (
      page: PptxGenJS.Slide,
      title: string,
      x = 0.68,
      y = 0.95,
      w = 5.2,
    ) => {
      page.addText(title, {
        x,
        y,
        w,
        h: 0.72,
        color: C.white,
        fontFace: "Graphik LC",
        bold: true,
        fontSize: 22,
        fit: "shrink",
        margin: 0,
        breakLine: false,
      });
    };

    const addCodeVisual = (page: PptxGenJS.Slide, x: number, y: number) => {
      page.addShape(pptx.ShapeType.roundRect, {
        x,
        y,
        w: 4.1,
        h: 3.25,
        rectRadius: 0.08,
        line: { color: C.line, transparency: 25 },
        fill: { color: C.panel, transparency: 10 },
      });
      [0.25, 0.55, 0.85, 1.15, 1.45, 1.75, 2.05].forEach((offset, i) => {
        page.addShape(pptx.ShapeType.line, {
          x: x + 0.35,
          y: y + offset,
          w: 2.55 + (i % 3) * 0.35,
          h: 0,
          line: { color: i % 2 ? C.soft : C.blue, transparency: 25, width: 1.3 },
        });
      });
      page.addShape(pptx.ShapeType.arc, {
        x: x + 2.85,
        y: y + 2.0,
        w: 0.9,
        h: 0.9,
        line: { color: C.blue, width: 2 },
        fill: { color: C.blue, transparency: 88 },
      });
    };

    const brandImages = await Promise.all(
      Object.entries(brandSlideAssets).map(async ([key, src]) => {
        try {
          return [key, await loadImageDataUrl(src)] as const;
        } catch {
          return [key, ""] as const;
        }
      }),
    ).then((entries) => Object.fromEntries(entries) as Record<keyof typeof brandSlideAssets, string>);

    currentSlides.forEach((slide, index) => {
      const page = pptx.addSlide();
      C = getSlideTheme(index) === "light" ? lightPalette : darkPalette;
      addBase(page, index + 1);

      if (slide.kind === "cover") {
        const coverTitle =
          slide.title === "Титульный слайд" ? presentationType : slide.title;
        addBlueRule(page, 0.68, 1.35);
        page.addText("Project", {
          x: 0.68,
          y: 1.72,
          w: 4.4,
          h: 0.48,
          color: C.white,
          bold: true,
          fontSize: 34,
          margin: 0,
        });
        page.addText(coverTitle, {
          x: 0.68,
          y: 2.25,
          w: 5.5,
          h: 0.58,
          color: C.blue,
          bold: true,
          fontSize: 36,
          margin: 0,
          fit: "shrink",
        });
        page.addText(slide.body ?? "Building digital products that drive results", {
          x: 0.7,
          y: 3.08,
          w: 4.5,
          h: 0.46,
          color: "FFFFFF",
          fontSize: 15,
          margin: 0,
          breakLine: false,
        });
        page.addShape(pptx.ShapeType.rect, {
          x: 6.15,
          y: 0,
          w: 7.18,
          h: 7.5,
          line: { color: C.dark, transparency: 100 },
          fill: { color: C.blue, transparency: 88 },
        });
        if (brandImages.coverPhoto) {
          page.addImage({
            data: brandImages.coverPhoto,
            x: 6.15,
            y: 0,
            w: 7.18,
            h: 7.5,
            transparency: 18,
          });
          page.addShape(pptx.ShapeType.rect, {
            x: 6.15,
            y: 0,
            w: 7.18,
            h: 7.5,
            line: { color: C.dark, transparency: 100 },
            fill: { color: C.dark, transparency: 32 },
          });
        } else {
          addCodeVisual(page, 7.45, 1.65);
        }
        page.addShape(pptx.ShapeType.line, {
          x: 7.1,
          y: 5.55,
          w: 3.8,
          h: -0.4,
          line: { color: C.blue, transparency: 30, width: 1.5 },
        });
        page.addText(new Date().toLocaleDateString("ru-RU"), {
          x: 0.7,
          y: 6.28,
          w: 2.2,
          h: 0.22,
          color: C.muted,
          fontSize: 8,
          margin: 0,
        });
        return;
      }

      if (slide.kind === "final") {
        page.addText("05", {
          x: 0.68,
          y: 1.35,
          w: 1.7,
          h: 0.78,
          color: C.blue,
          bold: true,
          fontSize: 46,
          margin: 0,
        });
        addTitle(page, slide.title, 0.72, 2.22, 4.6);
        page.addText(slide.body ?? "Готовы ответить на ваши вопросы", {
          x: 0.75,
          y: 3.08,
          w: 4.4,
          h: 0.35,
          color: C.muted,
          fontSize: 12,
          margin: 0,
        });
        page.addShape(pptx.ShapeType.roundRect, {
          x: 8.35,
          y: 1.78,
          w: 2.4,
          h: 1.65,
          rectRadius: 0.14,
          line: { color: C.line, transparency: 20 },
          fill: { color: C.panel2, transparency: 5 },
        });
        if (brandImages.questionIllustration) {
          page.addImage({
            data: brandImages.questionIllustration,
            x: 7.45,
            y: 1.12,
            w: 4.45,
            h: 4.25,
          });
        } else {
          page.addText("?", {
            x: 8.85,
            y: 1.9,
            w: 1.35,
            h: 1.2,
            color: C.blue,
            fontSize: 70,
            bold: true,
            align: "center",
            margin: 0,
          });
        }
        page.addShape(pptx.ShapeType.rect, {
          x: 0,
          y: 6.3,
          w: 13.33,
          h: 1.2,
          line: { color: C.blue },
          fill: { color: C.blue },
        });
        page.addText("Let's build something great together!", {
          x: 2.75,
          y: 6.52,
          w: 4.3,
          h: 0.45,
          color: C.white,
          fontSize: 18,
          bold: true,
          margin: 0,
        });
        page.addText("Questions? We're here to help.", {
          x: 10.2,
          y: 6.68,
          w: 2.1,
          h: 0.24,
          color: "CFDEFF",
          fontSize: 9,
          align: "right",
          margin: 0,
        });
        return;
      }

      addBlueRule(page);
      addTitle(page, slide.title);

      if (slide.kind === "agenda") {
        const agendaItems =
          slide.bullets && slide.bullets.length > 0
            ? slide.bullets
            : currentSlides.slice(2, -1).map((item) => item.title);
        agendaItems.forEach((item, itemIndex) => {
          page.addText(String(itemIndex + 1).padStart(2, "0"), {
            x: 0.72,
            y: 1.85 + itemIndex * 0.43,
            w: 0.35,
            h: 0.2,
            color: C.blue,
            bold: true,
            fontSize: 10,
            margin: 0,
          });
          page.addText(item, {
            x: 1.22,
            y: 1.82 + itemIndex * 0.43,
            w: 7,
            h: 0.28,
            color: C.white,
            fontSize: 12,
            margin: 0,
          });
        });
        if (brandImages.agendaPhoto) {
          page.addImage({
            data: brandImages.agendaPhoto,
            x: 6.45,
            y: 0,
            w: 6.88,
            h: 7.5,
            transparency: 10,
          });
          page.addShape(pptx.ShapeType.rect, {
            x: 5.75,
            y: 0,
            w: 7.58,
            h: 7.5,
            line: { color: C.dark, transparency: 100 },
            fill: { color: C.dark, transparency: 40 },
          });
        } else {
          addCodeVisual(page, 7.55, 1.1);
        }
        return;
      }

      if (slide.kind === "metrics") {
        page.addText("Key metrics", {
          x: 0.72,
          y: 1.9,
          w: 3.1,
          h: 0.3,
          color: C.muted,
          fontSize: 12,
          margin: 0,
        });
        const metricItems = slide.bullets ?? defaultBulletsByKind.metrics;
        ["85%", "60%", "4.7"].forEach((value, metricIndex) => {
          page.addText(value, {
            x: 0.72 + metricIndex * 2.4,
            y: 2.45,
            w: 1.65,
            h: 0.42,
            color: C.blue,
            bold: true,
            fontSize: 28,
            margin: 0,
          });
          page.addText(
            metricItems[metricIndex]?.replace(/^[-+0-9.%\s]+/, "") ??
              ["положительная оценка", "использование функции", "средняя оценка"][
                metricIndex
              ],
            {
              x: 0.74 + metricIndex * 2.4,
              y: 3.0,
              w: 1.5,
              h: 0.42,
              color: C.white,
              fontSize: 9,
              margin: 0,
            },
          );
        });
        const points = [
          [0.8, 5.95],
          [1.6, 5.55],
          [2.4, 5.82],
          [3.25, 5.35],
          [4.1, 5.18],
          [4.95, 5.48],
          [5.8, 5.22],
          [6.7, 5.38],
          [7.55, 5.05],
          [8.45, 5.16],
          [9.28, 4.82],
          [10.15, 4.98],
        ];
        points.slice(0, -1).forEach((point, pointIndex) => {
          const next = points[pointIndex + 1];
          page.addShape(pptx.ShapeType.line, {
            x: point[0],
            y: point[1],
            w: next[0] - point[0],
            h: next[1] - point[1],
            line: { color: C.blue, width: 2 },
          });
        });
        page.addShape(pptx.ShapeType.rect, {
          x: 0.72,
          y: 6.08,
          w: 9.6,
          h: 0.45,
          line: { color: C.blue, transparency: 100 },
          fill: { color: C.blue, transparency: 78 },
        });
        if (brandImages.metricsIllustration) {
          page.addImage({
            data: brandImages.metricsIllustration,
            x: 8.15,
            y: 1.45,
            w: 3.75,
            h: 3.35,
          });
        }
        return;
      }

      const assignedImage = getSlideImage(index);
      const layout = getSlideLayout(slideLayouts, index);

      if (slide.kind === "image" && assignedImage) {
        page.addText("02", {
          x: 0.7,
          y: 1.75,
          w: 1.3,
          h: 0.62,
          color: C.blue,
          bold: true,
          fontSize: 42,
          margin: 0,
        });
        page.addText(slide.title, {
          x: 0.76,
          y: 2.62,
          w: 4.2,
          h: 0.48,
          color: C.white,
          bold: true,
          fontSize: 22,
          margin: 0,
        });
        page.addImage({
          data: assignedImage.dataUrl,
          x: 6.8 + layout.imageX / 90,
          y: 1.05 + layout.imageY / 90,
          w: 4.9 * (layout.imageScale / 100),
          h: 5.5 * (layout.imageScale / 100),
          transparency: layout.imageTransparency,
        });
        return;
      }

      (slide.bullets ?? defaultBulletsByKind.content).slice(0, 4).forEach(
        (item, itemIndex) => {
          page.addShape(pptx.ShapeType.roundRect, {
            x: 0.72,
            y: 2.0 + itemIndex * 0.74,
            w: 0.2,
            h: 0.2,
            rectRadius: 0.03,
            line: { color: C.blue },
            fill: { color: C.blue },
          });
          page.addText(item, {
            x: 1.05,
            y: 1.95 + itemIndex * 0.74,
            w: 4.8,
            h: 0.28,
            color: C.white,
            bold: true,
            fontSize: 12,
            margin: 0,
          });
        },
      );
      page.addText(
        manualText.trim()
          ? manualText.trim().slice(0, 190)
          : textFile
            ? `Материалы взяты из файла: ${textFile.name}`
            : slide.body ?? "Здесь будет краткое содержание из загруженных материалов.",
        {
          x: 0.72,
          y: 4.7,
          w: 4.7,
          h: 0.85,
          color: C.muted,
          fontSize: 12,
          breakLine: false,
          fit: "shrink",
          margin: 0,
        },
      );
      if (brandImages.rocketIllustration && index % 3 === 2) {
        page.addImage({
          data: brandImages.rocketIllustration,
          x: 7.6,
          y: 1.25,
          w: 3.95,
          h: 4.9,
        });
      } else if (brandImages.contentPhoto) {
        page.addImage({
          data: brandImages.contentPhoto,
          x: 6.35,
          y: 0.85,
          w: 5.4,
          h: 5.8,
          transparency: 14,
        });
        page.addShape(pptx.ShapeType.rect, {
          x: 6.15,
          y: 0.85,
          w: 5.65,
          h: 5.8,
          line: { color: C.dark, transparency: 100 },
          fill: { color: C.dark, transparency: 40 },
        });
      } else {
        addCodeVisual(page, 7.25, 2.05);
      }
    });

    await pptx.writeFile({ fileName: "umbrella-presentation.pptx" });
  };

  return (
    <main className="min-h-screen bg-umbrella-bg font-sans text-umbrella-ink">
      <Header />

      <div className="grid min-h-[calc(100vh-72px)] gap-5 px-7 py-6 xl:grid-cols-[420px_1fr]">
        <aside className="rounded-lg border border-umbrella-line bg-white p-6 shadow-panel">
          <div>
            <h1 className="font-display text-2xl font-bold tracking-normal">
              Создать презентацию
            </h1>
            <p className="mt-2 max-w-[320px] text-sm leading-6 text-umbrella-muted">
              Загрузите материалы, на основе которых мы создадим презентацию
            </p>
          </div>

          <section className="mt-5">
            <div className="mb-4 grid grid-cols-2 border-b border-umbrella-line text-sm font-semibold">
              <button
                className={`px-4 pb-3 ${
                  textSource === "text"
                    ? "border-b-2 border-umbrella-accent text-umbrella-accent"
                    : "text-umbrella-muted"
                }`}
                onClick={() => setTextSource("text")}
              >
                Текст
              </button>
              <button
                className={`px-4 pb-3 ${
                  textSource === "file"
                    ? "border-b-2 border-umbrella-accent text-umbrella-accent"
                    : "text-umbrella-muted"
                }`}
                onClick={() => setTextSource("file")}
              >
                Файл
              </button>
            </div>

            {textSource === "text" ? (
              <textarea
                className="min-h-[160px] w-full resize-none rounded-lg border border-umbrella-line bg-white p-4 text-sm leading-6 outline-none transition focus:border-umbrella-accent focus:ring-4 focus:ring-[#0050FF]/10"
                placeholder="Вставьте текст, из которого нужно собрать презентацию"
                value={manualText}
                onChange={(event) => setManualText(event.target.value)}
              />
            ) : (
              <>
                <input
                  ref={textFileRef}
                  className="hidden"
                  type="file"
                  accept={sourceFileAccept}
                  onChange={(event: ChangeEvent<HTMLInputElement>) =>
                    handleTextFile(event.target.files?.[0])
                  }
                />
                <div
                  className="rounded-lg border border-dashed border-[#9DBBFF] bg-umbrella-soft p-6 text-center"
                  onDragOver={(event) => event.preventDefault()}
                  onDrop={onDropText}
                >
                  <FileText className="mx-auto h-11 w-11 text-umbrella-accent" />
                  <p className="mt-3 text-sm font-semibold">
                    Перетащите файл сюда
                  </p>
                  <p className="mt-1 text-sm text-umbrella-muted">
                    или выберите на компьютере
                  </p>
                  <p className="mt-2 text-xs text-umbrella-muted">
                    Поддерживаются: .docx, .txt, .pdf, .pptx, .png, .jpg, .webp
                  </p>
                  <button
                    className="mt-3 rounded-md border border-[#9DBBFF] bg-white px-4 py-2 text-sm font-semibold text-umbrella-accent transition hover:bg-umbrella-blueSoft"
                    onClick={() => textFileRef.current?.click()}
                  >
                    Выбрать файл
                  </button>
                </div>
              </>
            )}

            {textFile && (
              <div className="mt-4 flex items-center gap-3 rounded-lg border border-umbrella-line bg-white p-3">
                <FileText className="h-6 w-6 shrink-0 text-umbrella-accent" />
                <div className="min-w-0 flex-1">
                  <p className="truncate text-sm font-semibold">{textFile.name}</p>
                  <p className="text-xs text-umbrella-muted">
                    {formatSize(textFile.size)}
                  </p>
                </div>
                <button
                  className="rounded-md p-1 text-umbrella-muted transition hover:bg-[#F1F2F7] hover:text-umbrella-ink"
                  aria-label="Удалить файл"
                  onClick={() => setTextFile(null)}
                >
                  <X className="h-4 w-4" />
                </button>
              </div>
            )}
          </section>

          <section className="mt-5">
            <h2 className="text-sm font-bold">Изображения и скриншоты</h2>
            <p className="mt-1 text-sm leading-5 text-umbrella-muted">
              Добавьте скриншоты или изображения, которые нужно использовать в
              презентации
            </p>
            <input
              ref={imageFileRef}
              className="hidden"
              type="file"
              accept="image/png,image/jpeg,image/webp"
              multiple
              onChange={(event) => void handleImageFiles(event.target.files ?? undefined)}
            />
            <div
              className="mt-3 rounded-lg border border-dashed border-[#9DBBFF] bg-white px-4 py-4 text-center"
              onDragOver={(event) => event.preventDefault()}
              onDrop={onDropImages}
            >
              <button
                className="inline-flex items-center gap-2 text-sm font-semibold text-umbrella-accent"
                onClick={() => imageFileRef.current?.click()}
              >
                <ImageIcon className="h-7 w-7" />
                Загрузить изображения
              </button>
              <p className="mt-1 text-xs text-umbrella-muted">
                PNG, JPG, WEBP до 10 файлов
              </p>
            </div>

            {images.length > 0 && (
              <div className="mt-4">
                <p className="mb-2 text-sm font-bold">Изображения</p>
                <div className="grid grid-cols-2 gap-3">
                  {images.map((image) => (
                    <ImageMockupCard
                      key={image.id}
                      image={image}
                      mockup={getImageMockup(image)}
                      onMockupChange={(mockup) => updateImageMockup(image.id, mockup)}
                      onRemove={() => removeImage(image.id)}
                    />
                  ))}
                </div>
              </div>
            )}
          </section>

          <section className="mt-5">
            <button
              className="flex w-full items-center justify-between rounded-lg border border-umbrella-line bg-white px-4 py-3 text-left text-sm font-bold transition hover:bg-[#F7F7FB]"
              type="button"
              onClick={() => setIsStandardLibraryOpen((current) => !current)}
            >
              <span className="inline-flex items-center gap-2">
                <Library className="h-4 w-4 text-umbrella-accent" />
                Библиотека изображений
              </span>
              <ChevronDown
                className={`h-4 w-4 text-umbrella-muted transition ${
                  isStandardLibraryOpen ? "rotate-180" : ""
                }`}
              />
            </button>
            {isStandardLibraryOpen && (
              <div className="mt-3 grid grid-cols-2 gap-3">
                {standardImages.map((image) => (
                  <LibraryImageCard key={image.id} image={image} />
                ))}
              </div>
            )}
          </section>

          <section className="mt-5">
            <button
              className="flex w-full items-center justify-between rounded-lg border border-umbrella-line bg-white px-4 py-3 text-left text-sm font-bold transition hover:bg-[#F7F7FB]"
              type="button"
              onClick={() => setIsChartLibraryOpen((current) => !current)}
            >
              <span className="inline-flex items-center gap-2">
                <BarChart3 className="h-4 w-4 text-umbrella-accent" />
                Диаграммы
              </span>
              <ChevronDown
                className={`h-4 w-4 text-umbrella-muted transition ${
                  isChartLibraryOpen ? "rotate-180" : ""
                }`}
              />
            </button>
            {isChartLibraryOpen && (
              <div className="mt-3 rounded-lg border border-umbrella-line bg-[#FBFBFE] p-3">
                <textarea
                  className="min-h-[92px] w-full resize-none rounded-lg border border-umbrella-line bg-white px-3 py-2 text-sm leading-5 outline-none transition focus:border-umbrella-accent focus:ring-4 focus:ring-[#0050FF]/10"
                  placeholder="Опишите, какую диаграмму нужно получить"
                  value={chartDescription}
                  onChange={(event) => setChartDescription(event.target.value)}
                />
                <div className="mt-3 grid grid-cols-[1fr_auto] gap-2">
                  <select
                    className="h-10 rounded-lg border border-umbrella-line bg-white px-3 text-sm font-semibold outline-none"
                    value={chartType}
                    onChange={(event) => setChartType(event.target.value as ChartType)}
                  >
                    {chartTypeOptions.map((option) => (
                      <option key={option.value} value={option.value}>
                        {option.label}
                      </option>
                    ))}
                  </select>
                  <button
                    className="rounded-lg bg-umbrella-accent px-4 text-sm font-bold text-white transition hover:bg-[#0042D6]"
                    type="button"
                    onClick={generateChart}
                  >
                    Создать
                  </button>
                </div>
                <input
                  ref={chartFileRef}
                  className="hidden"
                  type="file"
                  accept=".csv,.tsv,.xlsx,.xls"
                  onChange={(event) => handleChartTableFile(event.target.files?.[0])}
                />
                <input
                  ref={chartImageRef}
                  className="hidden"
                  type="file"
                  accept="image/png,image/jpeg,image/webp"
                  multiple
                  onChange={(event) => void handleChartImageFiles(event.target.files ?? undefined)}
                />
                <div className="mt-3 grid grid-cols-2 gap-2">
                  <button
                    className="rounded-md border border-umbrella-line bg-white px-3 py-2 text-xs font-bold transition hover:bg-umbrella-soft"
                    type="button"
                    onClick={() => chartFileRef.current?.click()}
                  >
                    Загрузить таблицу
                  </button>
                  <button
                    className="rounded-md border border-umbrella-line bg-white px-3 py-2 text-xs font-bold transition hover:bg-umbrella-soft"
                    type="button"
                    onClick={() => chartImageRef.current?.click()}
                  >
                    Вставить картинку
                  </button>
                </div>
                {chartImages.length > 0 && (
                  <div className="mt-3 grid grid-cols-2 gap-3">
                    {chartImages.map((image) => (
                      <LibraryImageCard
                        key={image.id}
                        image={image}
                        onRemove={() => removeImage(image.id)}
                      />
                    ))}
                  </div>
                )}
              </div>
            )}
          </section>

          <section className="mt-5 space-y-4">
            <SelectField
              label="Тип презентации"
              value={presentationType}
              options={presentationTypes}
              onChange={(value) => setPresentationType(value as PresentationType)}
            />
            <SelectField
              label="Стиль"
              value={style}
              options={styles}
              onChange={(value) => setStyle(value as PresentationStyle)}
              icon={<LogoMark compact />}
            />
            <ThemeSegmentedControl
              label="Тема презентации"
              value={deckTheme}
              onChange={setDeckTheme}
            />
            <SelectField
              label="Количество слайдов"
              value={slideCount}
              options={slideCounts}
              onChange={(value) => setSlideCount(value as SlideCount)}
            />
          </section>

          <button
            className="mt-5 flex h-12 w-full items-center justify-center gap-2 rounded-lg bg-umbrella-accent px-4 font-bold text-white shadow-lg shadow-[#0050FF]/20 transition hover:bg-[#0042D6] disabled:cursor-wait disabled:opacity-70"
            disabled={isGenerating}
            onClick={() => void createPresentation()}
          >
            <Sparkles className={`h-5 w-5 ${isGenerating ? "animate-pulse" : ""}`} />
            {isGenerating ? "Генерируем..." : "Создать презентацию"}
          </button>
          {generationError && (
            <p className="mt-3 text-xs leading-4 text-umbrella-muted">
              {generationError}
            </p>
          )}
        </aside>

        <section className="flex min-w-0 flex-col rounded-lg border border-umbrella-line bg-white shadow-panel">
          <div className="flex items-center justify-between gap-4 px-6 py-5">
            <div className="flex items-center gap-3">
              <h2 className="font-display text-xl font-bold">Предпросмотр презентации</h2>
              {hasGenerated && (
                <span className="rounded-lg bg-[#F0F1F6] px-3 py-1 text-sm font-semibold">
                  {currentSlides.length} слайдов
                </span>
              )}
              {generationPrompt && (
                <span className="hidden max-w-[280px] truncate rounded-lg bg-umbrella-soft px-3 py-1 text-sm font-semibold text-umbrella-accent xl:inline">
                  {presentationPrompts[presentationType].structure.join(" → ")}
                </span>
              )}
            </div>
            <div className="flex items-center gap-2">
              <button
                className="inline-flex items-center gap-2 rounded-lg border border-umbrella-line bg-white px-4 py-2 text-sm font-semibold transition hover:bg-[#F7F7FB] disabled:cursor-not-allowed disabled:opacity-45"
                type="button"
                disabled={undoStack.length === 0}
                onClick={revertLastStep}
              >
                <Undo2 className="h-4 w-4" />
                Отменить
              </button>
              <button className="inline-flex items-center gap-2 rounded-lg border border-umbrella-line bg-white px-4 py-2 text-sm font-semibold transition hover:bg-[#F7F7FB]">
                <Pencil className="h-4 w-4" />
                Редактировать структуру
              </button>
            </div>
          </div>

          <div className="flex-1 border-y border-umbrella-line px-6 py-5">
            {hasGenerated ? (
              <div className="grid gap-6 lg:grid-cols-2 2xl:grid-cols-3">
                {currentSlides.map((slide, index) => (
                  <SlideCard
                    key={`${slide.title}-${index}`}
                    slide={slide}
                    index={index}
                    assignedImage={getSlideImage(index)}
                    imageMockup={getSlideImageMockup(index)}
                    layout={getSlideLayout(slideLayouts, index)}
                    slides={currentSlides}
                    theme={getSlideTheme(index)}
                    onThemeToggle={() => toggleSlideTheme(index)}
                    onImageDrop={(imageId) => assignImageToSlide(index, imageId)}
                    onOpen={() => setEditingSlideIndex(index)}
                    onAdd={() => addSlideAfter(index)}
                    onDelete={() => deleteSlideAt(index)}
                    onRegenerate={() => regenerateSlide(index)}
                  />
                ))}
              </div>
            ) : (
              <div className="flex min-h-[520px] items-center justify-center rounded-lg border border-dashed border-umbrella-line bg-[#FBFBFE] text-center">
                <div>
                  <Plus className="mx-auto h-9 w-9 text-umbrella-accent" />
                  <p className="mt-3 font-semibold">Предпросмотр появится здесь</p>
                  <p className="mt-1 text-sm text-umbrella-muted">
                    Загрузите материалы и создайте структуру
                  </p>
                </div>
              </div>
            )}
          </div>

          <div className="flex items-center justify-between px-6 py-4">
            <button className="rounded-lg border border-umbrella-line bg-white px-7 py-3 text-sm font-bold transition hover:bg-[#F7F7FB]">
              Назад
            </button>
            <button
              className="inline-flex items-center gap-2 rounded-lg border border-umbrella-line bg-white px-6 py-3 text-sm font-bold transition hover:border-[#9DBBFF] hover:bg-umbrella-soft"
              onClick={() => void downloadPptx()}
            >
              <Download className="h-5 w-5" />
              Скачать презентацию (.pptx)
            </button>
          </div>
        </section>
      </div>
      {editingSlideIndex !== null && currentSlides[editingSlideIndex] && (
        <SlideEditorModal
          slide={currentSlides[editingSlideIndex]}
          index={editingSlideIndex}
          theme={getSlideTheme(editingSlideIndex)}
          image={getSlideImage(editingSlideIndex)}
          imageMockup={getSlideImageMockup(editingSlideIndex)}
          layout={getSlideLayout(slideLayouts, editingSlideIndex)}
          versions={slideVersions[editingSlideIndex] ?? []}
          onClose={() => setEditingSlideIndex(null)}
          onChange={(patch) => updateSlideText(editingSlideIndex, patch)}
          onThemeChange={(theme) => updateSlideTheme(editingSlideIndex, theme)}
          onLayoutChange={(patch) => updateSlideLayout(editingSlideIndex, patch)}
          onLayoutChangeStart={commitSlideLayoutChange}
          onRestoreVersion={(version) => restoreSlideVersion(editingSlideIndex, version)}
          onPasteImage={() => void pasteImageToSlide(editingSlideIndex)}
          onAddSlide={() => addSlideAfter(editingSlideIndex)}
          onDeleteSlide={() => {
            deleteSlideAt(editingSlideIndex);
            setEditingSlideIndex(null);
          }}
          onRegenerate={() => regenerateSlide(editingSlideIndex)}
        />
      )}
    </main>
  );
}

function Header() {
  return (
    <header className="flex h-[72px] items-center justify-between border-b border-umbrella-line bg-white px-7">
      <div className="flex items-center gap-24">
        <div className="flex items-center gap-3">
          <LogoMark />
        </div>
        <nav className="hidden items-center gap-16 text-sm font-semibold text-umbrella-muted md:flex">
          <a href="#">Мои презентации</a>
          <a href="#">Шаблоны</a>
        </nav>
      </div>
      <div className="flex items-center gap-5">
        <button
          className="rounded-full p-1 text-umbrella-muted transition hover:bg-umbrella-soft hover:text-umbrella-ink"
          aria-label="Помощь"
        >
          <CircleHelp className="h-6 w-6" />
        </button>
        <div className="grid h-10 w-10 place-items-center rounded-full bg-umbrella-accent text-sm font-bold text-white">
          AK
        </div>
      </div>
    </header>
  );
}

function LogoMark({ compact = false }: { compact?: boolean }) {
  return (
    <span
      className={
        compact
          ? "block h-5 w-7 rounded bg-white p-0.5"
          : "block h-8 w-40"
      }
    >
      <img
        src={compact ? "/brand/uit-logo-short.svg" : "/brand/uit-logo-full.svg"}
        alt="Umbrella IT"
        className="h-full w-full object-contain object-left"
      />
    </span>
  );
}

function SelectField({
  label,
  value,
  options,
  onChange,
  icon,
}: {
  label: string;
  value: string;
  options: string[];
  onChange: (value: string) => void;
  icon?: ReactNode;
}) {
  return (
    <label className="block">
      <span className="mb-2 block text-sm font-bold">{label}</span>
      <span className="relative flex items-center">
        {icon && <span className="absolute left-3 z-10">{icon}</span>}
        <select
          className={`h-11 w-full appearance-none rounded-lg border border-umbrella-line bg-white pr-10 text-sm font-semibold outline-none transition focus:border-umbrella-accent focus:ring-4 focus:ring-[#0050FF]/10 ${
            icon ? "pl-10" : "pl-3"
          }`}
          value={value}
          onChange={(event) => onChange(event.target.value)}
        >
          {options.map((option) => (
            <option key={option} value={option}>
              {option}
            </option>
          ))}
        </select>
        <ChevronDown className="pointer-events-none absolute right-3 h-4 w-4 text-umbrella-muted" />
      </span>
    </label>
  );
}

function ThemeSegmentedControl({
  label,
  value,
  onChange,
}: {
  label: string;
  value: SlideTheme;
  onChange: (value: SlideTheme) => void;
}) {
  return (
    <div>
      <span className="mb-2 block text-sm font-bold">{label}</span>
      <div className="grid h-11 grid-cols-2 rounded-lg border border-umbrella-line bg-white p-1">
        {slideThemeOptions.map((option) => (
          <button
            key={option.value}
            className={`rounded-md text-sm font-semibold transition ${
              value === option.value
                ? "bg-umbrella-accent text-white shadow-sm"
                : "text-umbrella-muted hover:bg-umbrella-soft hover:text-umbrella-ink"
            }`}
            type="button"
            onClick={() => onChange(option.value)}
          >
            {option.label}
          </button>
        ))}
      </div>
    </div>
  );
}

function RangeControl({
  label,
  value,
  min,
  max,
  onChange,
  onChangeStart,
}: {
  label: string;
  value: number;
  min: number;
  max: number;
  onChange: (value: number) => void;
  onChangeStart: () => void;
}) {
  return (
    <label className="block">
      <span className="mb-1 flex items-center justify-between text-xs font-bold text-umbrella-muted">
        <span>{label}</span>
        <span>{value}</span>
      </span>
      <input
        className="w-full accent-[#0050FF]"
        type="range"
        min={min}
        max={max}
        value={value}
        onMouseDown={onChangeStart}
        onTouchStart={onChangeStart}
        onChange={(event) => onChange(Number(event.target.value))}
      />
    </label>
  );
}

function LibraryImageCard({
  image,
  onRemove,
}: {
  image: UploadedImage;
  onRemove?: () => void;
}) {
  return (
    <div
      className="relative cursor-grab overflow-hidden rounded-lg border border-umbrella-line bg-[#F2F3F8] p-2 transition hover:border-[#9DBBFF] active:cursor-grabbing"
      draggable
      onDragStart={(event) => {
        event.dataTransfer.setData("application/x-umbrella-image-id", image.id);
        event.dataTransfer.effectAllowed = "copy";
      }}
      title="Перетащите на слайд"
    >
      <div className="grid h-24 place-items-center rounded-md bg-white">
        <img
          src={image.dataUrl}
          alt={image.name}
          className="max-h-20 max-w-full object-contain"
        />
      </div>
      <p className="mt-2 truncate text-xs font-semibold text-umbrella-muted">
        {image.name}
      </p>
      <p className="mt-1 text-[11px] font-medium text-umbrella-muted">
        Перетащите на слайд
      </p>
      {onRemove && (
        <button
          className="absolute right-2 top-2 rounded-full bg-white p-1 text-umbrella-muted shadow transition hover:text-umbrella-ink"
          aria-label="Удалить"
          onClick={onRemove}
          type="button"
        >
          <X className="h-3.5 w-3.5" />
        </button>
      )}
    </div>
  );
}

function ImageMockupCard({
  image,
  mockup,
  onMockupChange,
  onRemove,
}: {
  image: UploadedImage;
  mockup: MockupChoice;
  onMockupChange: (mockup: MockupChoice) => void;
  onRemove: () => void;
}) {
  const option = getMockupOption(mockup);
  const Icon = option.icon;

  return (
    <div
      className="relative cursor-grab overflow-hidden rounded-lg border border-umbrella-line bg-[#F2F3F8] p-2 transition hover:border-[#9DBBFF] active:cursor-grabbing"
      draggable
      onDragStart={(event) => {
        event.dataTransfer.setData("application/x-umbrella-image-id", image.id);
        event.dataTransfer.effectAllowed = "copy";
      }}
      title="Перетащите на слайд"
    >
      <div className="grid h-24 place-items-center rounded-md bg-white">
        <MockupFrame image={image} mockup={mockup} compact />
      </div>
      <div className="mt-2 flex items-center justify-between gap-2">
        <p className="min-w-0 truncate text-xs font-semibold text-umbrella-muted">
          {image.name}
        </p>
        <label
          className="relative inline-flex shrink-0 items-center gap-1 text-xs font-bold text-umbrella-accent"
          onClick={(event) => event.stopPropagation()}
        >
          <Icon className="h-3.5 w-3.5" />
          <select
            className="max-w-[98px] cursor-pointer appearance-none bg-transparent pr-4 font-bold outline-none"
            value={mockup}
            onChange={(event) => onMockupChange(event.target.value as MockupChoice)}
            onMouseDown={(event) => event.stopPropagation()}
          >
            {mockupOptions.map((item) => (
              <option key={item.value} value={item.value}>
                {item.label}
              </option>
            ))}
          </select>
          <ChevronDown className="pointer-events-none absolute right-0 h-3 w-3" />
        </label>
      </div>
      <p className="mt-1 text-[11px] font-medium text-umbrella-muted">
        Перетащите на слайд
      </p>
      <button
        className="absolute right-2 top-2 rounded-full bg-white p-1 text-umbrella-muted shadow transition hover:text-umbrella-ink"
        aria-label="Удалить изображение"
        onClick={onRemove}
        type="button"
      >
        <X className="h-3.5 w-3.5" />
      </button>
    </div>
  );
}

function SlideCard({
  slide,
  index,
  assignedImage,
  imageMockup,
  layout,
  slides,
  theme,
  onThemeToggle,
  onImageDrop,
  onOpen,
  onAdd,
  onDelete,
  onRegenerate,
}: {
  slide: Slide;
  index: number;
  assignedImage?: UploadedImage;
  imageMockup?: MockupChoice;
  layout: SlideLayout;
  slides: Slide[];
  theme: SlideTheme;
  onThemeToggle: () => void;
  onImageDrop: (imageId: string) => void;
  onOpen: () => void;
  onAdd: () => void;
  onDelete: () => void;
  onRegenerate: () => void;
}) {
  const isLight = theme === "light";
  const [isDragOver, setIsDragOver] = useState(false);

  const handleDrop = (event: DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    setIsDragOver(false);

    const imageId = event.dataTransfer.getData("application/x-umbrella-image-id");
    if (imageId) {
      onImageDrop(imageId);
    }
  };

  return (
    <article>
      <div
        className="relative"
        onDragEnter={(event) => {
          if (event.dataTransfer.types.includes("application/x-umbrella-image-id")) {
            setIsDragOver(true);
          }
        }}
        onDragOver={(event) => {
          if (event.dataTransfer.types.includes("application/x-umbrella-image-id")) {
            event.preventDefault();
            event.dataTransfer.dropEffect = "copy";
            setIsDragOver(true);
          }
        }}
        onDragLeave={(event) => {
          if (!event.currentTarget.contains(event.relatedTarget as Node | null)) {
            setIsDragOver(false);
          }
        }}
        onDrop={handleDrop}
      >
        <button
          className="slide-aspect block w-full overflow-hidden rounded-md border border-umbrella-line bg-white text-left shadow-card transition hover:border-[#9DBBFF] hover:shadow-panel"
          type="button"
          onClick={onOpen}
          aria-label={`Открыть слайд ${index + 1}`}
        >
          <SlidePreview
            slide={slide}
            image={assignedImage}
            imageMockup={imageMockup}
            layout={layout}
            slides={slides}
            theme={theme}
          />
        </button>
        {isDragOver && (
          <div className="pointer-events-none absolute inset-0 z-20 grid place-items-center rounded-md border-2 border-dashed border-umbrella-accent bg-[#0050FF]/15 p-4 text-center">
            <span className="rounded-full bg-white px-4 py-2 text-sm font-bold text-umbrella-accent shadow-lg">
              Отпустите, чтобы заменить изображение
            </span>
          </div>
        )}
        <button
          className={`absolute right-3 top-3 grid h-9 w-9 place-items-center rounded-full border shadow-lg transition duration-300 hover:scale-105 ${
            isLight
              ? "border-[#D8DAE2] bg-white text-[#F6A400]"
              : "border-white/20 bg-[#07111C] text-umbrella-blueSoft"
          }`}
          type="button"
          aria-label={isLight ? "Сделать слайд тёмным" : "Сделать слайд светлым"}
          title={isLight ? "Тёмная тема" : "Светлая тема"}
          onClick={onThemeToggle}
        >
          {isLight ? (
            <Sun key="sun" className="h-4 w-4 animate-theme-pop" />
          ) : (
            <Moon key="moon" className="h-4 w-4 animate-theme-pop" />
          )}
        </button>
      </div>
      <div className="mt-3 flex items-center justify-between gap-3">
        <div className="min-w-0 flex-1">
          <p className="truncate text-sm font-semibold text-umbrella-muted">
            {index + 1}. {slide.title}
          </p>
          {assignedImage && (
            <p className="mt-0.5 truncate text-xs font-semibold text-umbrella-accent">
              {assignedImage.name}
            </p>
          )}
        </div>
        <div className="flex shrink-0 items-center gap-1">
          <button
            className="grid h-8 w-8 place-items-center rounded-md border border-umbrella-line bg-white text-umbrella-muted transition hover:bg-umbrella-soft hover:text-umbrella-accent"
            type="button"
            aria-label="Увеличить и редактировать"
            onClick={onOpen}
            title="Увеличить и редактировать"
          >
            <Maximize2 className="h-4 w-4" />
          </button>
          <button
            className="grid h-8 w-8 place-items-center rounded-md border border-umbrella-line bg-white text-umbrella-muted transition hover:bg-umbrella-soft hover:text-umbrella-accent"
            type="button"
            aria-label="Добавить слайд"
            onClick={onAdd}
            title="Добавить слайд"
          >
            <Plus className="h-4 w-4" />
          </button>
          <button
            className="grid h-8 w-8 place-items-center rounded-md border border-umbrella-line bg-white text-umbrella-muted transition hover:bg-umbrella-soft hover:text-umbrella-accent"
            type="button"
            aria-label="Перегенерировать слайд"
            onClick={onRegenerate}
            title="Перегенерировать"
          >
            <RotateCcw className="h-4 w-4" />
          </button>
          <button
            className="grid h-8 w-8 place-items-center rounded-md border border-umbrella-line bg-white text-umbrella-muted transition hover:bg-[#FFF1F1] hover:text-[#D23B3B]"
            type="button"
            aria-label="Удалить слайд"
            onClick={onDelete}
            title="Удалить слайд"
          >
            <Trash2 className="h-4 w-4" />
          </button>
        </div>
      </div>
    </article>
  );
}

function SlideEditorModal({
  slide,
  index,
  theme,
  image,
  imageMockup,
  layout,
  versions,
  onClose,
  onChange,
  onThemeChange,
  onLayoutChange,
  onLayoutChangeStart,
  onRestoreVersion,
  onPasteImage,
  onAddSlide,
  onDeleteSlide,
  onRegenerate,
}: {
  slide: Slide;
  index: number;
  theme: SlideTheme;
  image?: UploadedImage;
  imageMockup?: MockupChoice;
  layout: SlideLayout;
  versions: Slide[];
  onClose: () => void;
  onChange: (patch: Partial<Pick<Slide, "title" | "body" | "bullets">>) => void;
  onThemeChange: (theme: SlideTheme) => void;
  onLayoutChange: (patch: Partial<SlideLayout>) => void;
  onLayoutChangeStart: () => void;
  onRestoreVersion: (version: Slide) => void;
  onPasteImage: () => void;
  onAddSlide: () => void;
  onDeleteSlide: () => void;
  onRegenerate: () => void;
}) {
  const bullets = slide.bullets ?? defaultBulletsByKind[slide.kind];

  const updateBullet = (bulletIndex: number, value: string) => {
    const nextBullets = [...bullets];
    nextBullets[bulletIndex] = value;
    onChange({ bullets: nextBullets });
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-[#07111C]/70 p-6 backdrop-blur-sm">
      <div className="grid max-h-[92vh] w-full max-w-6xl grid-cols-[1.35fr_420px] overflow-hidden rounded-lg bg-white shadow-2xl">
        <div className="bg-[#111827] p-6">
          <div className="mb-4 flex items-center justify-between text-white">
            <div>
              <p className="text-sm text-white/60">Слайд {index + 1}</p>
              <h3 className="font-display text-2xl font-bold">{slide.title}</h3>
            </div>
            <button
              className="rounded-md bg-white/10 px-3 py-2 text-sm font-semibold transition hover:bg-white/15"
              type="button"
              onClick={onClose}
            >
              Закрыть
            </button>
          </div>
          <div className="slide-aspect overflow-hidden rounded-lg border border-white/10 bg-white">
            <SlidePreview
              slide={slide}
              image={image}
              imageMockup={imageMockup}
              layout={layout}
              slides={[slide]}
              theme={theme}
            />
          </div>
        </div>

        <div className="overflow-y-auto p-6">
          <div className="flex items-center justify-between gap-3">
            <h3 className="font-display text-xl font-bold">Редактировать текст</h3>
            <div className="flex items-center gap-1">
              <button
                className="grid h-9 w-9 place-items-center rounded-lg border border-umbrella-line bg-white text-umbrella-muted transition hover:bg-umbrella-soft hover:text-umbrella-accent"
                type="button"
                onClick={onPasteImage}
                title="Вставить картинку из буфера"
                aria-label="Вставить картинку из буфера"
              >
                <Clipboard className="h-4 w-4" />
              </button>
              <button
                className="grid h-9 w-9 place-items-center rounded-lg border border-umbrella-line bg-white text-umbrella-muted transition hover:bg-umbrella-soft hover:text-umbrella-accent"
                type="button"
                onClick={onAddSlide}
                title="Добавить слайд"
                aria-label="Добавить слайд"
              >
                <Plus className="h-4 w-4" />
              </button>
              <button
                className="grid h-9 w-9 place-items-center rounded-lg border border-umbrella-line bg-white text-umbrella-muted transition hover:bg-umbrella-soft hover:text-umbrella-accent"
                type="button"
                onClick={onRegenerate}
                title="Перегенерировать"
                aria-label="Перегенерировать"
              >
                <RotateCcw className="h-4 w-4" />
              </button>
              <button
                className="grid h-9 w-9 place-items-center rounded-lg border border-umbrella-line bg-white text-umbrella-muted transition hover:bg-[#FFF1F1] hover:text-[#D23B3B]"
                type="button"
                onClick={onDeleteSlide}
                title="Удалить слайд"
                aria-label="Удалить слайд"
              >
                <Trash2 className="h-4 w-4" />
              </button>
            </div>
          </div>

          <div className="mt-5">
            <ThemeSegmentedControl
              label="Тема слайда"
              value={theme}
              onChange={onThemeChange}
            />
          </div>

          <div className="mt-5 rounded-lg border border-umbrella-line bg-[#FBFBFE] p-4">
            <div className="flex items-center justify-between gap-3">
              <h4 className="text-sm font-bold">Макет</h4>
              <button
                className="text-xs font-bold text-umbrella-accent"
                type="button"
                onClick={() => onLayoutChange(defaultSlideLayout)}
              >
                Сбросить
              </button>
            </div>
            <div className="mt-3 grid gap-3">
              <RangeControl
                label="Размер изображения"
                value={layout.imageScale}
                min={60}
                max={300}
                onChangeStart={onLayoutChangeStart}
                onChange={(value) => onLayoutChange({ imageScale: value })}
              />
              <RangeControl
                label="Прозрачность изображения"
                value={layout.imageTransparency}
                min={0}
                max={90}
                onChangeStart={onLayoutChangeStart}
                onChange={(value) => onLayoutChange({ imageTransparency: value })}
              />
              <div className="grid grid-cols-2 gap-3">
                <RangeControl
                  label="Изображение X"
                  value={layout.imageX}
                  min={-80}
                  max={80}
                  onChangeStart={onLayoutChangeStart}
                  onChange={(value) => onLayoutChange({ imageX: value })}
                />
                <RangeControl
                  label="Изображение Y"
                  value={layout.imageY}
                  min={-80}
                  max={80}
                  onChangeStart={onLayoutChangeStart}
                  onChange={(value) => onLayoutChange({ imageY: value })}
                />
                <RangeControl
                  label="Текст X"
                  value={layout.textX}
                  min={-80}
                  max={80}
                  onChangeStart={onLayoutChangeStart}
                  onChange={(value) => onLayoutChange({ textX: value })}
                />
                <RangeControl
                  label="Текст Y"
                  value={layout.textY}
                  min={-80}
                  max={80}
                  onChangeStart={onLayoutChangeStart}
                  onChange={(value) => onLayoutChange({ textY: value })}
                />
              </div>
            </div>
          </div>

          {versions.length > 0 && (
            <div className="mt-5 rounded-lg border border-umbrella-line bg-white p-4">
              <div className="flex items-center gap-2">
                <History className="h-4 w-4 text-umbrella-accent" />
                <h4 className="text-sm font-bold">История версий</h4>
              </div>
              <div className="mt-3 space-y-2">
                {versions.map((version, versionIndex) => (
                  <button
                    key={`${version.title}-${versionIndex}`}
                    className="w-full rounded-md border border-umbrella-line px-3 py-2 text-left text-xs font-semibold transition hover:border-[#9DBBFF] hover:bg-umbrella-soft"
                    type="button"
                    onClick={() => onRestoreVersion(version)}
                  >
                    Версия {versions.length - versionIndex}: {version.title}
                  </button>
                ))}
              </div>
            </div>
          )}

          <label className="mt-5 block">
            <span className="mb-2 block text-sm font-bold">Заголовок</span>
            <input
              className="h-11 w-full rounded-lg border border-umbrella-line px-3 text-sm font-semibold outline-none transition focus:border-umbrella-accent focus:ring-4 focus:ring-[#0050FF]/10"
              value={slide.title}
              onChange={(event) => onChange({ title: event.target.value })}
            />
          </label>

          <label className="mt-4 block">
            <span className="mb-2 block text-sm font-bold">Описание</span>
            <textarea
              className="min-h-[96px] w-full resize-none rounded-lg border border-umbrella-line px-3 py-2 text-sm leading-5 outline-none transition focus:border-umbrella-accent focus:ring-4 focus:ring-[#0050FF]/10"
              value={slide.body ?? ""}
              onChange={(event) => onChange({ body: event.target.value })}
            />
          </label>

          <div className="mt-4">
            <span className="mb-2 block text-sm font-bold">Пункты</span>
            <div className="space-y-2">
              {bullets.slice(0, 5).map((bullet, bulletIndex) => (
                <input
                  key={bulletIndex}
                  className="h-10 w-full rounded-lg border border-umbrella-line px-3 text-sm outline-none transition focus:border-umbrella-accent focus:ring-4 focus:ring-[#0050FF]/10"
                  value={bullet}
                  onChange={(event) => updateBullet(bulletIndex, event.target.value)}
                />
              ))}
            </div>
          </div>

          <button
            className="mt-6 h-11 w-full rounded-lg bg-umbrella-accent px-4 text-sm font-bold text-white transition hover:bg-[#0042D6]"
            type="button"
            onClick={onClose}
          >
            Готово
          </button>
        </div>
      </div>
    </div>
  );
}

function MockupFrame({
  image,
  mockup,
  compact = false,
}: {
  image: UploadedImage;
  mockup: MockupChoice;
  compact?: boolean;
}) {
  if (mockup === "none") {
    return (
      <img
        src={image.dataUrl}
        alt={image.name}
        className={`rounded-md object-contain shadow-card ${
          compact ? "h-20 w-full" : "max-h-[190px] w-36"
        }`}
      />
    );
  }

  if (mockup === "iphone") {
    return (
      <div
        className={`relative overflow-hidden rounded-[20px] border-[5px] border-[#1E2028] bg-[#1E2028] shadow-card ${
          compact ? "h-20 w-11" : "h-full max-h-[190px] w-24"
        }`}
      >
        <span className="absolute left-1/2 top-1 z-10 h-1.5 w-6 -translate-x-1/2 rounded-full bg-[#1E2028]" />
        <img
          src={image.dataUrl}
          alt={image.name}
          className="h-full w-full rounded-[14px] object-cover"
        />
      </div>
    );
  }

  if (mockup === "macbook") {
    return (
      <div className={compact ? "w-28" : "w-36"}>
        <div className="overflow-hidden rounded-t-md border-[4px] border-[#232631] bg-[#232631] shadow-card">
          <img
            src={image.dataUrl}
            alt={image.name}
            className={`${compact ? "h-14" : "h-24"} w-full object-cover`}
          />
        </div>
        <div className="mx-auto h-2 w-[112%] -translate-x-[5%] rounded-b-lg bg-[#D8DAE2]" />
      </div>
    );
  }

  return (
    <div
      className={`overflow-hidden rounded-md border border-[#D8DAE2] bg-white shadow-card ${
        compact ? "w-28" : "w-36"
      }`}
    >
      <div className="flex h-4 items-center gap-1 border-b border-[#E5E7EF] bg-[#F7F7FB] px-2">
        <span className="h-1.5 w-1.5 rounded-full bg-[#FF6B6B]" />
        <span className="h-1.5 w-1.5 rounded-full bg-[#F6C85F]" />
        <span className="h-1.5 w-1.5 rounded-full bg-[#5BCB7A]" />
      </div>
      <img
        src={image.dataUrl}
        alt={image.name}
        className={`${compact ? "h-14" : "h-24"} w-full object-cover`}
      />
    </div>
  );
}

function SlidePreview({
  slide,
  image,
  imageMockup,
  layout,
  slides,
  theme,
}: {
  slide: Slide;
  image?: UploadedImage;
  imageMockup?: MockupChoice;
  layout: SlideLayout;
  slides: Slide[];
  theme: SlideTheme;
}) {
  const isLight = theme === "light";
  const previewBg = isLight ? "bg-[#F7F7F7]" : "bg-[#07111C]";
  const previewText = isLight ? "text-[#282A32]" : "text-white";
  const secondaryText = isLight ? "text-[#6B6D75]" : "text-white";
  const panelBg = isLight ? "bg-white" : "bg-white/5";
  const imageOpacity = isLight ? "opacity-20" : "opacity-35";

  if (slide.kind === "cover" || slide.kind === "final") {
    return (
      <div className={`relative h-full overflow-hidden p-5 ${previewBg} ${previewText}`}>
        <img
          src={
            slide.kind === "cover"
              ? brandSlideAssets.coverPhoto
              : brandSlideAssets.questionIllustration
          }
          alt=""
          className={`absolute inset-y-0 right-0 h-full w-1/2 object-cover ${imageOpacity}`}
        />
        <div className="absolute -right-8 top-3 h-40 w-40 rounded-full border-[34px] border-umbrella-accent opacity-70" />
        <LogoMark compact />
        <div className="mt-8 max-w-[70%]">
          {slide.kind === "cover" ? (
            <>
              <p className={isLight ? "text-[10px] text-umbrella-accent" : "text-[10px] text-umbrella-blueSoft"}>Демо проекта</p>
              <h3 className="mt-2 font-display text-xl font-bold leading-snug">
                Umbrella Deck Builder
              </h3>
              <p className={isLight ? "mt-6 text-[10px] text-umbrella-muted" : "mt-6 text-[10px] text-umbrella-blueSoft"}>Июнь 2026</p>
            </>
          ) : (
            <>
              <h3 className="mt-8 font-display text-2xl font-bold">Спасибо!</h3>
              <p className={isLight ? "mt-3 text-xs text-umbrella-muted" : "mt-3 text-xs text-umbrella-soft"}>
                Готовы ответить на ваши вопросы
              </p>
            </>
          )}
        </div>
      </div>
    );
  }

  if (slide.kind === "agenda") {
    const agendaItems =
      slide.bullets && slide.bullets.length > 0
        ? slide.bullets
        : slides.slice(2, -1).map((item) => item.title);
    return (
      <div className={`relative h-full overflow-hidden p-5 ${previewBg} ${previewText}`}>
        <img
          src={brandSlideAssets.agendaPhoto}
          alt=""
          className={`absolute inset-y-0 right-0 h-full w-1/2 object-cover ${imageOpacity}`}
        />
        <div className="absolute right-0 top-0 h-full w-1/2 bg-[radial-gradient(circle_at_center,#0050FF33,transparent_55%)]" />
        <div className="h-1 w-12 bg-umbrella-accent" />
        <h3 className="mt-4 font-display text-xl font-bold">{slide.title}</h3>
        <div className="mt-5 space-y-3">
          {agendaItems.slice(0, 6).map((item, index) => (
            <div key={`${item}-${index}`} className="flex items-center gap-4">
              <span className="w-6 text-[10px] font-bold text-umbrella-accent">
                {String(index + 1).padStart(2, "0")}
              </span>
              <span className={`text-xs font-medium ${secondaryText}`}>{item}</span>
            </div>
          ))}
        </div>
      </div>
    );
  }

  if (slide.kind === "metrics") {
    const metricItems = slide.bullets ?? defaultBulletsByKind.metrics;
    return (
      <div className={`relative h-full overflow-hidden p-5 ${previewBg} ${previewText}`}>
        <img
          src={brandSlideAssets.metricsIllustration}
          alt=""
          className="absolute bottom-5 right-8 h-36 w-36 object-contain opacity-90"
        />
        <div className="h-1 w-12 bg-umbrella-accent" />
        <h3 className="mt-4 font-display text-xl font-bold">{slide.title}</h3>
        <div className="mt-8 grid grid-cols-3 gap-4">
          {["85%", "60%", "4.7"].map((value, index) => (
            <div key={value} className={`rounded-md p-4 text-left ${panelBg}`}>
              <p className="text-2xl font-black text-umbrella-accent">{value}</p>
              <p className={`mt-2 text-[10px] leading-4 ${secondaryText}`}>
                {
                  metricItems[index]?.replace(/^[-+0-9.%\s]+/, "") ??
                  [
                    "положительная оценка",
                    "использование функции",
                    "средняя оценка",
                  ][index]
                }
              </p>
            </div>
          ))}
        </div>
        <div className="absolute bottom-6 left-5 right-5 h-12 border-b border-umbrella-accent/50 bg-gradient-to-t from-[#0050FF66] to-transparent" />
      </div>
    );
  }

  return (
    <div className={`relative grid h-full grid-cols-[1fr_auto] gap-4 overflow-hidden p-5 ${previewBg} ${previewText}`}>
      <img
        src={
          slide.kind === "image"
            ? brandSlideAssets.discussionIllustration
            : brandSlideAssets.contentPhoto
        }
        alt=""
        className={`absolute inset-y-0 right-0 h-full w-1/2 object-cover ${isLight ? "opacity-15" : "opacity-25"}`}
      />
      <div className="absolute right-4 top-6 h-28 w-28 rounded-full border-[18px] border-umbrella-accent/50" />
      <div
        className="relative z-10"
        style={{
          transform: `translate(${layout.textX}px, ${layout.textY}px)`,
        }}
      >
        <div className="h-1 w-12 bg-umbrella-accent" />
        <h3 className="mt-4 font-display text-xl font-bold">{slide.title}</h3>
        <div className="mt-7 space-y-3">
          {(slide.bullets ?? defaultBulletsByKind.content).slice(0, 3).map(
            (item) => (
              <div key={item} className="flex items-center gap-2">
                <span className="grid h-5 w-5 place-items-center rounded-full bg-umbrella-blueSoft text-[10px] font-bold text-umbrella-accent">
                  ✓
                </span>
                <span className={`text-xs font-medium ${secondaryText}`}>{item}</span>
              </div>
            ),
          )}
        </div>
      </div>
      {slide.kind === "image" && image ? (
        <div
          className="relative z-10 flex h-full w-36 items-center justify-center"
          style={{
            transform: `translate(${layout.imageX}px, ${layout.imageY}px) scale(${layout.imageScale / 100})`,
            opacity: (100 - layout.imageTransparency) / 100,
          }}
        >
          <MockupFrame
            image={image}
            mockup={imageMockup ?? "none"}
          />
        </div>
      ) : (
        <div className="relative z-10 mt-16 hidden h-16 w-16 rounded-full bg-umbrella-accent/30 md:block" />
      )}
    </div>
  );
}
