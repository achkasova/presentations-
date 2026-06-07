# Umbrella Deck Builder

MVP веб-приложения для генерации презентаций из текста, файлов и изображений в корпоративном стиле Umbrella.

## Что уже есть

- Next.js + TypeScript + Tailwind CSS.
- Загрузка текстового файла: `.docx`, `.txt`, `.pdf`.
- Ручная вставка текста.
- Загрузка до 10 изображений: PNG, JPG, WEBP.
- Миниатюры изображений с удалением.
- Настройки типа презентации, стиля и количества слайдов.
- Мок-генерация структуры слайдов по типу презентации.
- Предпросмотр слайдов в корпоративном стиле.
- Реальная генерация `.pptx` через `pptxgenjs`.

## Запуск локально

```bash
npm install
npm run dev
```

После запуска откройте:

```text
http://localhost:3000
```

## GitHub

```bash
git init
git add .
git commit -m "Initial Umbrella Deck Builder MVP"
git branch -M main
git remote add origin <URL_ВАШЕГО_GITHUB_РЕПОЗИТОРИЯ>
git push -u origin main
```

## Vercel

1. Создайте проект на Vercel.
2. Подключите GitHub-репозиторий.
3. Framework Preset: Next.js.
4. Build Command: `npm run build`.
5. Install Command: `npm install`.
6. Deploy.

## Следующий этап: OpenAI API

- Добавить серверный route handler, например `app/api/generate/route.ts`.
- Передавать в API извлеченный текст и метаданные загруженных файлов.
- Хранить ключ в `.env.local`: `OPENAI_API_KEY=...`.
- Заменить мок-структуру на JSON-ответ модели.
- Добавить обработку ошибок, состояния загрузки и валидацию JSON-структуры.
