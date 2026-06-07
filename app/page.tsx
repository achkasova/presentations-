"use client";

import {
  ChevronDown,
  Check,
  CircleHelp,
  Download,
  FileText,
  Image as ImageIcon,
  Laptop,
  Maximize2,
  Monitor,
  Moon,
  Pencil,
  Plus,
  RotateCcw,
  Sparkles,
  Smartphone,
  Sun,
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

type Slide = {
  title: string;
  kind: "cover" | "agenda" | "content" | "metrics" | "image" | "final";
  body?: string;
  bullets?: string[];
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
  value: MockupType;
  icon: typeof Smartphone;
}> = [
  { label: "iPhone", value: "iphone", icon: Smartphone },
  { label: "Browser", value: "browser", icon: Monitor },
  { label: "Macbook", value: "macbook", icon: Laptop },
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

const presentationFormatRequirements = [
  "Формат 16:9, расчетная база 1920x1080.",
  "Все слайды строятся на единой 12-column grid с внешними полями 80-100 px по бокам, 60-80 px сверху и 50-70 px снизу.",
  "Темная палитра: основной фон #020814, второй фон #07111E, карточки #0B1526, границы rgba(255,255,255,0.06).",
  "Акцент #005CFF, светлый акцент #3A7BFF, основной текст #FFFFFF, вторичный текст #B6C0D1.",
  "Шрифт должен быть Inter-подобным. Заголовки крупные и bold, без декоративной перегрузки.",
  "Cover: 45% текст и 55% изображение, структура: лого, синий штрих, заголовок, подзаголовок, дата.",
  "Agenda: 40% список и 60% фото, номера пунктов синим.",
  "Section Intro: крупный номер 80-100 px, заголовок 48-56 px, описание 18-22 px.",
  "KPI: карточки около 280x220 px, padding 24 px, число 48-60 px bold, описание 16 px.",
  "Screenshot slide: сетка 40/60, слева заголовок, описание и список, справа скриншот.",
  "Поддерживаемые мастер-слайды: Cover, Agenda, Section Intro, Overview, KPI, Feature + Screenshot, Technology Stack, Team, Timeline, Next Steps, Q&A, Thank You.",
];

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
    `Требования к формату: ${presentationFormatRequirements.join(" ")}`,
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

const acceptedTextFormats = [".docx", ".txt", ".pdf"];
const acceptedImageFormats = ["image/png", "image/jpeg", "image/webp"];

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

const getMockupOption = (type: MockupType) =>
  mockupOptions.find((option) => option.value === type) ?? mockupOptions[1];

const getFirstWord = (text: string) => text.trim().split(/\s+/)[0] || text;

const getPptxRuleWidth = (title: string) =>
  Math.max(0.42, Math.min(1.9, getFirstWord(title).length * 0.16));

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

export default function Home() {
  const [textSource, setTextSource] = useState<TextSource>("file");
  const [manualText, setManualText] = useState("");
  const [textFile, setTextFile] = useState<UploadedFile | null>(null);
  const [images, setImages] = useState<UploadedImage[]>([]);
  const [presentationType, setPresentationType] =
    useState<PresentationType>("Демо проекта");
  const [style, setStyle] = useState<PresentationStyle>("Umbrella корпоративный");
  const [slideCount, setSlideCount] = useState<SlideCount>("8–10");
  const [deckTheme, setDeckTheme] = useState<SlideTheme>("dark");
  const [slideThemes, setSlideThemes] = useState<Record<number, SlideTheme>>({});
  const [autoMockups, setAutoMockups] = useState(true);
  const [isMockupPanelOpen, setIsMockupPanelOpen] = useState(false);
  const [imageMockups, setImageMockups] = useState<Record<string, MockupType>>({});
  const [slides, setSlides] = useState<Slide[]>(demoSlides.map(hydrateSlide));
  const [generationPrompt, setGenerationPrompt] = useState("");
  const [hasGenerated, setHasGenerated] = useState(false);
  const [editingSlideIndex, setEditingSlideIndex] = useState<number | null>(null);
  const textFileRef = useRef<HTMLInputElement | null>(null);
  const imageFileRef = useRef<HTMLInputElement | null>(null);

  const materializeSlides = (baseSlides: Slide[]) => {
    if (images.length === 0) {
      return baseSlides.map((slide) =>
        slide.kind === "image"
          ? hydrateSlide({ ...slide, kind: "content" as const })
          : hydrateSlide(slide),
      );
    }

    if (!baseSlides.some((slide) => slide.kind === "image")) {
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
    [images.length, slides],
  );

  const handleTextFile = (file?: File) => {
    if (!file) return;
    const lowerName = file.name.toLowerCase();
    const isAllowed = acceptedTextFormats.some((format) =>
      lowerName.endsWith(format),
    );

    if (!isAllowed) return;
    setTextFile({ name: file.name, size: file.size, type: file.type });
  };

  const handleImageFiles = async (fileList?: FileList | File[]) => {
    if (!fileList) return;

    const incoming = Array.from(fileList)
      .filter((file) => acceptedImageFormats.includes(file.type))
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

  const onDropText = (event: DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    handleTextFile(event.dataTransfer.files[0]);
  };

  const onDropImages = (event: DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    void handleImageFiles(event.dataTransfer.files);
  };

  const createPresentation = () => {
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

    setGenerationPrompt(nextPrompt);
    setSlides(nextSlides);
    setSlideThemes(
      Object.fromEntries(nextSlides.map((_, index) => [index, deckTheme])),
    );
    setHasGenerated(true);
  };

  const getSlideTheme = (index: number): SlideTheme =>
    slideThemes[index] ?? deckTheme;

  const updateSlideTheme = (index: number, theme: SlideTheme) => {
    setSlideThemes((current) => ({ ...current, [index]: theme }));
  };

  const toggleSlideTheme = (index: number) => {
    updateSlideTheme(index, getSlideTheme(index) === "dark" ? "light" : "dark");
  };

  const getImageMockup = (image: UploadedImage): MockupType =>
    autoMockups ? inferImageMockup(image) : imageMockups[image.id] ?? inferImageMockup(image);

  const updateImageMockup = (imageId: string, mockup: MockupType) => {
    setImageMockups((current) => ({ ...current, [imageId]: mockup }));
  };

  const removeImage = (imageId: string) => {
    setImages((current) => current.filter((item) => item.id !== imageId));
    setImageMockups((current) => {
      const next = { ...current };
      delete next[imageId];
      return next;
    });
  };

  const replaceSlideAt = (index: number, nextSlide: Slide) => {
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

  const downloadPptx = async () => {
    const pptx = new PptxGenJS();
    pptx.layout = "LAYOUT_WIDE";
    pptx.author = "Umbrella Deck Builder";
    pptx.subject = `${presentationType} — ${style}`;
    pptx.title = "Umbrella Presentation";
    pptx.company = "Umbrella";
    pptx.theme = {
      headFontFace: "Inter",
      bodyFontFace: "Inter",
    };

    const layout = {
      slideW: 13.33,
      slideH: 7.5,
      marginX: 0.68,
      marginTop: 0.5,
      marginBottom: 0.42,
      gutter: 0.22,
      colW: 0.91,
      cardGap: 0.22,
    };

    const darkPalette = {
      dark: "020814",
      dark2: "07111E",
      panel: "0B1526",
      panel2: "07111E",
      blue: "005CFF",
      blue2: "3A7BFF",
      soft: "B6C0D1",
      white: "FFFFFF",
      muted: "B6C0D1",
      line: "FFFFFF",
    };
    const lightPalette = {
      dark: "F7F7F7",
      dark2: "FFFFFF",
      panel: "FFFFFF",
      panel2: "F2F6FF",
      blue: "005CFF",
      blue2: "3A7BFF",
      soft: "005CFF",
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
        w: layout.slideW,
        h: layout.slideH,
        line: { color: C.dark, transparency: 100 },
        fill: { color: C.dark },
      });
      page.addShape(pptx.ShapeType.rect, {
        x: 0,
        y: 0,
        w: layout.slideW,
        h: layout.slideH,
        line: { color: C.dark, transparency: 100 },
        fill: { color: C.dark2, transparency: 40 },
      });
      page.addText("Umbrella IT", {
        x: layout.marginX,
        y: 0.38,
        w: 1.35,
        h: 0.24,
        color: C.white,
        fontFace: "Inter",
        bold: true,
        fontSize: 9,
        margin: 0,
      });
      page.addShape(pptx.ShapeType.arc, {
        x: 0.49,
        y: 0.37,
        w: 0.18,
        h: 0.18,
        line: { color: C.blue, width: 1 },
        fill: { color: C.blue },
      });
      page.addShape(pptx.ShapeType.line, {
        x: layout.marginX,
        y: 6.94,
        w: 0.34,
        h: 0,
        line: { color: C.line, transparency: 94, width: 0.5 },
      });
      page.addText(String(pageNumber), {
        x: layout.marginX + 0.02,
        y: 6.86,
        w: 0.28,
        h: 0.18,
        color: C.muted,
        fontSize: 6,
        margin: 0,
      });
    };

    const addBlueRule = (
      page: PptxGenJS.Slide,
      title: string,
      x = layout.marginX,
      y = 0.86,
    ) => {
      page.addShape(pptx.ShapeType.rect, {
        x,
        y,
        w: getPptxRuleWidth(title),
        h: 0.04,
        line: { color: C.blue },
        fill: { color: C.blue },
      });
    };

    const addTitle = (
      page: PptxGenJS.Slide,
      title: string,
      x = layout.marginX,
      y = 0.95,
      w = 5.4,
    ) => {
      page.addText(title, {
        x,
        y,
        w,
        h: 0.72,
        color: C.white,
        fontFace: "Inter",
        bold: true,
        fontSize: 30,
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
        addBlueRule(page, coverTitle, layout.marginX, 1.58);
        page.addText("Project", {
          x: layout.marginX,
          y: 1.82,
          w: 4.4,
          h: 0.58,
          color: C.white,
          bold: true,
          fontSize: 42,
          margin: 0,
        });
        page.addText(coverTitle, {
          x: layout.marginX,
          y: 2.48,
          w: 4.95,
          h: 0.78,
          color: C.blue,
          bold: true,
          fontSize: 46,
          margin: 0,
          fit: "shrink",
        });
        page.addText(slide.body ?? "Building digital products that drive results", {
          x: layout.marginX,
          y: 3.45,
          w: 4.6,
          h: 0.46,
          color: C.muted,
          fontSize: 18,
          margin: 0,
          breakLine: false,
        });
        page.addShape(pptx.ShapeType.rect, {
          x: 6.0,
          y: 0,
          w: 7.33,
          h: layout.slideH,
          line: { color: C.dark, transparency: 100 },
          fill: { color: C.dark2 },
        });
        if (brandImages.coverPhoto) {
          page.addImage({
            data: brandImages.coverPhoto,
            x: 6.0,
            y: 0,
            w: 7.33,
            h: layout.slideH,
            transparency: 18,
          });
          page.addShape(pptx.ShapeType.rect, {
            x: 6.0,
            y: 0,
            w: 7.33,
            h: layout.slideH,
            line: { color: C.dark, transparency: 100 },
            fill: { color: C.dark, transparency: 32 },
          });
        } else {
          addCodeVisual(page, 7.45, 1.65);
        }
        page.addText(new Date().toLocaleDateString("ru-RU"), {
          x: layout.marginX,
          y: 6.28,
          w: 2.2,
          h: 0.22,
          color: C.muted,
          fontSize: 12,
          margin: 0,
        });
        return;
      }

      if (slide.kind === "final") {
        page.addText(String(index + 1).padStart(2, "0"), {
          x: layout.marginX,
          y: 1.35,
          w: 1.7,
          h: 0.78,
          color: C.blue,
          bold: true,
          fontSize: 68,
          margin: 0,
        });
        addTitle(page, slide.title, layout.marginX, 2.35, 4.8);
        page.addText(slide.body ?? "Готовы ответить на ваши вопросы", {
          x: layout.marginX,
          y: 3.28,
          w: 4.4,
          h: 0.35,
          color: C.muted,
          fontSize: 16,
          margin: 0,
        });
        page.addShape(pptx.ShapeType.roundRect, {
          x: 8.35,
          y: 1.78,
          w: 2.4,
          h: 1.65,
          rectRadius: 0.14,
          line: { color: C.line, transparency: 94 },
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
          w: layout.slideW,
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

      addBlueRule(page, slide.title, layout.marginX, 0.9);
      addTitle(page, slide.title);

      if (slide.kind === "agenda") {
        const agendaItems =
          slide.bullets && slide.bullets.length > 0
            ? slide.bullets
            : currentSlides.slice(2, -1).map((item) => item.title);
        agendaItems.forEach((item, itemIndex) => {
          page.addText(String(itemIndex + 1).padStart(2, "0"), {
            x: layout.marginX,
            y: 1.85 + itemIndex * 0.52,
            w: 0.42,
            h: 0.24,
            color: C.blue,
            bold: true,
            fontSize: 12,
            margin: 0,
          });
          page.addText(item, {
            x: 1.24,
            y: 1.81 + itemIndex * 0.52,
            w: 4.2,
            h: 0.3,
            color: C.white,
            fontSize: 14,
            margin: 0,
          });
        });
        if (brandImages.agendaPhoto) {
          page.addImage({
            data: brandImages.agendaPhoto,
            x: 5.45,
            y: 0,
            w: 7.88,
            h: layout.slideH,
            transparency: 10,
          });
          page.addShape(pptx.ShapeType.rect, {
            x: 5.45,
            y: 0,
            w: 7.88,
            h: layout.slideH,
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
          x: layout.marginX,
          y: 1.9,
          w: 3.1,
          h: 0.3,
          color: C.muted,
          fontSize: 14,
          margin: 0,
        });
        const metricItems = slide.bullets ?? defaultBulletsByKind.metrics;
        ["85%", "60%", "4.7"].forEach((value, metricIndex) => {
          const cardX = layout.marginX + metricIndex * 2.22;
          page.addShape(pptx.ShapeType.roundRect, {
            x: cardX,
            y: 2.42,
            w: 1.95,
            h: 1.55,
            rectRadius: 0.08,
            line: { color: C.line, transparency: 94 },
            fill: { color: C.panel },
          });
          page.addText(value, {
            x: cardX + 0.2,
            y: 2.68,
            w: 1.45,
            h: 0.48,
            color: C.blue,
            bold: true,
            fontSize: 40,
            margin: 0,
          });
          page.addText(
            metricItems[metricIndex]?.replace(/^[-+0-9.%\s]+/, "") ??
              ["положительная оценка", "использование функции", "средняя оценка"][
                metricIndex
              ],
            {
              x: cardX + 0.2,
              y: 3.25,
              w: 1.5,
              h: 0.42,
              color: C.muted,
              fontSize: 12,
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
          x: layout.marginX,
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

      if (slide.kind === "image" && images[0]) {
        page.addText(String(index + 1).padStart(2, "0"), {
          x: layout.marginX,
          y: 1.45,
          w: 1.3,
          h: 0.62,
          color: C.blue,
          bold: true,
          fontSize: 68,
          margin: 0,
        });
        page.addText(slide.title, {
          x: layout.marginX,
          y: 2.58,
          w: 4.4,
          h: 0.62,
          color: C.white,
          bold: true,
          fontSize: 30,
          margin: 0,
        });
        page.addText(slide.body ?? "Ключевой экран и сценарий использования", {
          x: layout.marginX,
          y: 3.42,
          w: 4.35,
          h: 0.55,
          color: C.muted,
          fontSize: 15,
          fit: "shrink",
          margin: 0,
        });
        page.addShape(pptx.ShapeType.roundRect, {
          x: 5.65,
          y: 0.82,
          w: 6.75,
          h: 5.75,
          rectRadius: 0.1,
          line: { color: C.line, transparency: 94 },
          fill: { color: C.panel },
        });
        page.addImage({
          data: images[0].dataUrl,
          x: 5.92,
          y: 1.08,
          w: 6.2,
          h: 5.2,
        });
        return;
      }

      (slide.bullets ?? defaultBulletsByKind.content).slice(0, 4).forEach(
        (item, itemIndex) => {
          page.addShape(pptx.ShapeType.roundRect, {
            x: layout.marginX,
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
            fontSize: 14,
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
          x: layout.marginX,
          y: 4.7,
          w: 4.7,
          h: 0.85,
          color: C.muted,
          fontSize: 14,
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
                  accept=".docx,.txt,.pdf"
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
                    Поддерживаются: .docx, .txt, .pdf
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
                      onRemove={() => removeImage(image.id)}
                    />
                  ))}
                </div>

                <label className="mt-4 flex items-start gap-3 rounded-lg border border-umbrella-line bg-[#FBFBFE] p-3">
                  <span className="relative mt-0.5 grid h-5 w-5 shrink-0 place-items-center rounded border border-umbrella-line bg-white">
                    <input
                      className="peer absolute inset-0 opacity-0"
                      type="checkbox"
                      checked={autoMockups}
                      onChange={(event) => setAutoMockups(event.target.checked)}
                    />
                    <Check className="h-3.5 w-3.5 text-transparent peer-checked:text-umbrella-accent" />
                  </span>
                  <span className="min-w-0">
                    <span className="block text-sm font-semibold">
                      Автоматически подбирать мокапы
                    </span>
                    <span className="mt-1 block text-xs leading-4 text-umbrella-muted">
                      AI определит подходящий мокап для каждого изображения
                    </span>
                  </span>
                </label>

                <button
                  className="mt-3 text-sm font-semibold text-umbrella-accent transition hover:text-[#0042D6]"
                  type="button"
                  onClick={() => {
                    setIsMockupPanelOpen((current) => !current);
                    setAutoMockups(false);
                  }}
                >
                  Настроить вручную →
                </button>

                {isMockupPanelOpen && (
                  <div className="mt-3 space-y-3 rounded-lg border border-umbrella-line bg-white p-3">
                    {images.map((image) => (
                      <div key={image.id} className="grid gap-2">
                        <p className="truncate text-xs font-semibold text-umbrella-muted">
                          {image.name}
                        </p>
                        <div className="grid grid-cols-3 gap-2">
                          {mockupOptions.map((option) => {
                            const Icon = option.icon;
                            const isSelected = getImageMockup(image) === option.value;

                            return (
                              <button
                                key={option.value}
                                className={`grid h-16 place-items-center rounded-md border text-xs font-semibold transition ${
                                  isSelected
                                    ? "border-umbrella-accent bg-umbrella-soft text-umbrella-accent"
                                    : "border-umbrella-line bg-white text-umbrella-muted hover:border-[#9DBBFF] hover:text-umbrella-ink"
                                }`}
                                type="button"
                                onClick={() => updateImageMockup(image.id, option.value)}
                              >
                                <Icon className="h-5 w-5" />
                                <span>{option.label}</span>
                              </button>
                            );
                          })}
                        </div>
                      </div>
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
            className="mt-5 flex h-12 w-full items-center justify-center gap-2 rounded-lg bg-umbrella-accent px-4 font-bold text-white shadow-lg shadow-[#0050FF]/20 transition hover:bg-[#0042D6]"
            onClick={createPresentation}
          >
            <Sparkles className="h-5 w-5" />
            Создать презентацию
          </button>
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
            <button className="inline-flex items-center gap-2 rounded-lg border border-umbrella-line bg-white px-4 py-2 text-sm font-semibold transition hover:bg-[#F7F7FB]">
              <Pencil className="h-4 w-4" />
              Редактировать структуру
            </button>
          </div>

          <div className="flex-1 border-y border-umbrella-line px-6 py-5">
            {hasGenerated ? (
              <div className="grid gap-6 lg:grid-cols-2 2xl:grid-cols-3">
                {currentSlides.map((slide, index) => (
                  <SlideCard
                    key={`${slide.title}-${index}`}
                    slide={slide}
                    index={index}
                    images={images}
                    imageMockup={images[0] ? getImageMockup(images[0]) : undefined}
                    slides={currentSlides}
                    theme={getSlideTheme(index)}
                    onThemeToggle={() => toggleSlideTheme(index)}
                    onOpen={() => setEditingSlideIndex(index)}
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
          images={images}
          imageMockup={images[0] ? getImageMockup(images[0]) : undefined}
          onClose={() => setEditingSlideIndex(null)}
          onChange={(patch) => updateSlideText(editingSlideIndex, patch)}
          onThemeChange={(theme) => updateSlideTheme(editingSlideIndex, theme)}
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

function ImageMockupCard({
  image,
  mockup,
  onRemove,
}: {
  image: UploadedImage;
  mockup: MockupType;
  onRemove: () => void;
}) {
  const option = getMockupOption(mockup);
  const Icon = option.icon;

  return (
    <div className="relative overflow-hidden rounded-lg border border-umbrella-line bg-[#F2F3F8] p-2">
      <div className="grid h-24 place-items-center rounded-md bg-white">
        <MockupFrame image={image} mockup={mockup} compact />
      </div>
      <div className="mt-2 flex items-center justify-between gap-2">
        <p className="min-w-0 truncate text-xs font-semibold text-umbrella-muted">
          {image.name}
        </p>
        <span className="inline-flex shrink-0 items-center gap-1 text-xs font-bold text-umbrella-accent">
          <Icon className="h-3.5 w-3.5" />
          {option.label}
        </span>
      </div>
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
  images,
  imageMockup,
  slides,
  theme,
  onThemeToggle,
  onOpen,
  onRegenerate,
}: {
  slide: Slide;
  index: number;
  images: UploadedImage[];
  imageMockup?: MockupType;
  slides: Slide[];
  theme: SlideTheme;
  onThemeToggle: () => void;
  onOpen: () => void;
  onRegenerate: () => void;
}) {
  const isLight = theme === "light";

  return (
    <article>
      <div className="relative">
        <button
          className="slide-aspect block w-full overflow-hidden rounded-md border border-umbrella-line bg-white text-left shadow-card transition hover:border-[#9DBBFF] hover:shadow-panel"
          type="button"
          onClick={onOpen}
          aria-label={`Открыть слайд ${index + 1}`}
        >
          <SlidePreview
            slide={slide}
            images={images}
            imageMockup={imageMockup}
            slides={slides}
            theme={theme}
          />
        </button>
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
        <p className="min-w-0 flex-1 truncate text-sm font-semibold text-umbrella-muted">
          {index + 1}. {slide.title}
        </p>
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
            aria-label="Перегенерировать слайд"
            onClick={onRegenerate}
            title="Перегенерировать"
          >
            <RotateCcw className="h-4 w-4" />
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
  images,
  imageMockup,
  onClose,
  onChange,
  onThemeChange,
  onRegenerate,
}: {
  slide: Slide;
  index: number;
  theme: SlideTheme;
  images: UploadedImage[];
  imageMockup?: MockupType;
  onClose: () => void;
  onChange: (patch: Partial<Pick<Slide, "title" | "body" | "bullets">>) => void;
  onThemeChange: (theme: SlideTheme) => void;
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
              images={images}
              imageMockup={imageMockup}
              slides={[slide]}
              theme={theme}
            />
          </div>
        </div>

        <div className="overflow-y-auto p-6">
          <div className="flex items-center justify-between gap-3">
            <h3 className="font-display text-xl font-bold">Редактировать текст</h3>
            <button
              className="inline-flex items-center gap-2 rounded-lg border border-umbrella-line bg-white px-3 py-2 text-sm font-semibold text-umbrella-muted transition hover:bg-umbrella-soft hover:text-umbrella-accent"
              type="button"
              onClick={onRegenerate}
            >
              <RotateCcw className="h-4 w-4" />
              Перегенерировать
            </button>
          </div>

          <div className="mt-5">
            <ThemeSegmentedControl
              label="Тема слайда"
              value={theme}
              onChange={onThemeChange}
            />
          </div>

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
  mockup: MockupType;
  compact?: boolean;
}) {
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

function TitleWithAccent({
  title,
  className,
}: {
  title: string;
  className: string;
}) {
  const [firstWord = title, ...restWords] = title.trim().split(/\s+/);
  const restTitle = restWords.join(" ");

  return (
    <h3 className={className}>
      <span className="inline-flex flex-col items-start align-top">
        <span className="mb-3 h-1 w-full bg-[#005CFF]" />
        <span>{firstWord}</span>
      </span>
      {restTitle ? ` ${restTitle}` : ""}
    </h3>
  );
}

function SlidePreview({
  slide,
  images,
  imageMockup,
  slides,
  theme,
}: {
  slide: Slide;
  images: UploadedImage[];
  imageMockup?: MockupType;
  slides: Slide[];
  theme: SlideTheme;
}) {
  const isLight = theme === "light";
  const previewBg = isLight ? "bg-[#F7F7F7]" : "bg-[#020814]";
  const previewText = isLight ? "text-[#282A32]" : "text-white";
  const secondaryText = isLight ? "text-[#6B6D75]" : "text-[#B6C0D1]";
  const panelBg = isLight ? "bg-white" : "bg-[#0B1526]";
  const imageOpacity = isLight ? "opacity-20" : "opacity-35";

  if (slide.kind === "cover" || slide.kind === "final") {
    const coverTitle = slide.kind === "cover" ? "Umbrella Deck Builder" : "Спасибо!";

    return (
      <div className={`relative h-full overflow-hidden p-5 ${previewBg} ${previewText}`}>
        <img
          src={
            slide.kind === "cover"
              ? brandSlideAssets.coverPhoto
              : brandSlideAssets.questionIllustration
          }
          alt=""
          className={`absolute inset-y-0 right-0 h-full w-[55%] object-cover ${imageOpacity}`}
        />
        <LogoMark compact />
        <div className="mt-10 max-w-[45%]">
          {slide.kind === "cover" ? (
            <>
              <p className={isLight ? "mt-3 text-[10px] text-[#005CFF]" : "mt-3 text-[10px] text-[#3A7BFF]"}>Демо проекта</p>
              <TitleWithAccent
                title={coverTitle}
                className="mt-2 font-display text-2xl font-bold leading-snug"
              />
              <p className={isLight ? "mt-6 text-[10px] text-umbrella-muted" : "mt-6 text-[10px] text-[#B6C0D1]"}>Июнь 2026</p>
            </>
          ) : (
            <>
              <p className="text-5xl font-black text-[#005CFF]">12</p>
              <TitleWithAccent
                title={coverTitle}
                className="mt-4 font-display text-2xl font-bold leading-snug"
              />
              <p className={isLight ? "mt-3 text-xs text-umbrella-muted" : "mt-3 text-xs text-[#B6C0D1]"}>
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
          className={`absolute inset-y-0 right-0 h-full w-[60%] object-cover ${imageOpacity}`}
        />
        <div className="absolute right-0 top-0 h-full w-[60%] bg-[#020814]/45" />
        <TitleWithAccent
          title={slide.title}
          className="mt-5 font-display text-xl font-bold leading-snug"
        />
        <div className="mt-5 space-y-3">
          {agendaItems.slice(0, 6).map((item, index) => (
            <div key={`${item}-${index}`} className="flex items-center gap-4">
              <span className="w-6 text-[10px] font-bold text-[#005CFF]">
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
        <TitleWithAccent
          title={slide.title}
          className="mt-5 font-display text-xl font-bold leading-snug"
        />
        <div className="mt-8 grid grid-cols-3 gap-4">
          {["85%", "60%", "4.7"].map((value, index) => (
            <div key={value} className={`rounded-md border border-white/[0.06] p-4 text-left ${panelBg}`}>
              <p className="text-2xl font-black text-[#005CFF]">{value}</p>
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
        <div className="absolute bottom-6 left-5 right-5 h-12 border-b border-[#005CFF]/50 bg-gradient-to-t from-[#005CFF66] to-transparent" />
      </div>
    );
  }

  return (
    <div className={`relative grid h-full grid-cols-[40%_1fr] gap-4 overflow-hidden p-5 ${previewBg} ${previewText}`}>
      <img
        src={
          slide.kind === "image"
            ? brandSlideAssets.discussionIllustration
            : brandSlideAssets.contentPhoto
        }
        alt=""
        className={`absolute inset-y-0 right-0 h-full w-[60%] object-cover ${isLight ? "opacity-15" : "opacity-18"}`}
      />
      <div className="relative z-10">
        <TitleWithAccent
          title={slide.title}
          className="mt-5 font-display text-xl font-bold leading-snug"
        />
        <div className="mt-7 space-y-3">
          {(slide.bullets ?? defaultBulletsByKind.content).slice(0, 3).map(
            (item) => (
              <div key={item} className="flex items-center gap-2">
                <span className="grid h-5 w-5 place-items-center rounded-full bg-[#3A7BFF]/20 text-[10px] font-bold text-[#3A7BFF]">
                  ✓
                </span>
                <span className={`text-xs font-medium ${secondaryText}`}>{item}</span>
              </div>
            ),
          )}
        </div>
      </div>
      {slide.kind === "image" && images[0] ? (
        <div className="relative z-10 flex h-full items-center justify-center rounded-md border border-white/[0.06] bg-[#0B1526] p-3">
          <MockupFrame
            image={images[0]}
            mockup={imageMockup ?? inferImageMockup(images[0])}
          />
        </div>
      ) : (
        <div className="relative z-10 mt-16 hidden h-16 w-16 rounded-full bg-umbrella-accent/30 md:block" />
      )}
    </div>
  );
}
