import { NextResponse } from "next/server";

type SlideKind = "cover" | "agenda" | "content" | "metrics" | "image" | "final";

type GeneratedSlide = {
  title: string;
  kind: SlideKind;
  body?: string;
  bullets?: string[];
};

const allowedKinds = new Set<SlideKind>([
  "cover",
  "agenda",
  "content",
  "metrics",
  "image",
  "final",
]);

const normalizeSlides = (value: unknown): GeneratedSlide[] => {
  if (!Array.isArray(value)) return [];

  const slides: GeneratedSlide[] = [];

  value.forEach((item) => {
    if (!item || typeof item !== "object") return;

    const slide = item as Record<string, unknown>;
    const title = typeof slide.title === "string" ? slide.title.trim() : "";
    const kind = typeof slide.kind === "string" ? slide.kind : "";

    if (!title || !allowedKinds.has(kind as SlideKind)) return;

    slides.push({
      title,
      kind: kind as SlideKind,
      body: typeof slide.body === "string" ? slide.body : undefined,
      bullets: Array.isArray(slide.bullets)
        ? slide.bullets
            .filter((bullet): bullet is string => typeof bullet === "string")
            .map((bullet) => bullet.trim())
            .filter(Boolean)
            .slice(0, 6)
        : undefined,
    });
  });

  return slides;
};

const extractText = (response: unknown) => {
  if (!response || typeof response !== "object") return "";

  const data = response as {
    output_text?: unknown;
    output?: Array<{
      content?: Array<{ text?: unknown; type?: unknown }>;
    }>;
  };

  if (typeof data.output_text === "string") {
    return data.output_text;
  }

  return (
    data.output
      ?.flatMap((item) => item.content ?? [])
      .map((content) => (typeof content.text === "string" ? content.text : ""))
      .join("\n") ?? ""
  );
};

export async function POST(request: Request) {
  const apiKey = process.env.OPENAI_API_KEY;

  if (!apiKey) {
    return NextResponse.json(
      { error: "OPENAI_API_KEY is not configured" },
      { status: 500 },
    );
  }

  const body = (await request.json()) as {
    prompt?: unknown;
    fallbackSlides?: unknown;
  };
  const prompt = typeof body.prompt === "string" ? body.prompt : "";
  const fallbackSlides = normalizeSlides(body.fallbackSlides);

  if (!prompt.trim()) {
    return NextResponse.json({ error: "Prompt is required" }, { status: 400 });
  }

  const response = await fetch("https://api.openai.com/v1/responses", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${apiKey}`,
    },
    body: JSON.stringify({
      model: process.env.OPENAI_MODEL ?? "gpt-4.1-mini",
      instructions:
        "Ты создаёшь структуру презентации. Верни только валидный JSON-массив без markdown. Каждый элемент: title, kind, body, bullets. kind строго один из: cover, agenda, content, metrics, image, final.",
      input: prompt,
      max_output_tokens: 1800,
    }),
  });

  if (!response.ok) {
    const errorText = await response.text();

    return NextResponse.json(
      {
        error: "OpenAI request failed",
        details: errorText.slice(0, 500),
      },
      { status: response.status },
    );
  }

  const data = await response.json();
  const text = extractText(data).trim();

  try {
    const parsed = JSON.parse(text);
    const slides = normalizeSlides(parsed);

    if (slides.length > 0) {
      return NextResponse.json({ slides });
    }
  } catch {
    // Fall through to the local fallback below.
  }

  return NextResponse.json({
    slides: fallbackSlides,
    warning: "AI response was not valid JSON, fallback slides were used",
  });
}
