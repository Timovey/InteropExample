# Interop Example

Пример заполнения шаблона Word-документа из кода: в `ExampleTemplate.docx` метки вида `[TUB]` заменяются
значениями из словаря, результат сохраняется в `Result.docx`.

**Стек:** C# · .NET 6 · Open XML SDK (`DocumentFormat.OpenXml`)

## Как это работает

1. Шаблон копируется в `Result.docx`.
2. Open XML SDK находит абзацы с метками `[ИМЯ]` (регулярное выражение `\[\w+\]`).
3. Метки, для которых есть значение в словаре `marks`, заменяются текстом.

Open XML SDK работает напрямую с форматом `.docx`, поэтому установленный Microsoft Word для этого не нужен. В проекте
также подключён COM Interop Word — для сравнения подходов.

## Запуск

```bash
cd InteropExample
dotnet run
```
