# Contributing в OpenDesk

Спасибо за интерес к проекту! 🎉

## Как помочь
1. Форкни репозиторий и создай ветку: `git checkout -b feat/my-feature`
2. Установи зависимости: `npm install`
3. Запусти dev и тесты:
   ```bash
   npm run dev
   npm test
   ```
4. Следуй стилю: TypeScript strict, без `any` без нужды, маленькие компоненты.
5. Добавь/обнови тесты для `src/excel/formula.ts`, если меняешь формулы.
6. Коммит в стиле Conventional Commits:
   - `feat: добавить выравнивание в Word`
   - `fix: исправить SUM с пустыми ячейками`
   - `docs: обновить README`
7. Открой Pull Request с описанием «что/зачем/как проверено».

## Правила кода
- React 18 + function components + hooks.
- Никаких тяжёлых editor-зависимостей без обсуждения в issue.
- Формулы Excel — чистые функции, без DOM.
- Все пользовательские строки на русском + позже i18n.

## Проверка PR
- `npm test` — зелёный
- `npm run build` — собирается без ошибок TS
- Ручная проверка в Chrome/Edge/Firefox

## Code of Conduct
Будь вежлив. Не принимаем оскорбления, спам, закрытый код без лицензии.
По вопросам — открывай Discussion / Issue.
