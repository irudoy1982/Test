# Khalil Audit System

## Руководство разработчика

Версия документа: X3-dev.3
Владелец продукта: Ivan Rudoy

### 1. Назначение продукта

Khalil Audit System - веб-приложение для сбора сведений об ИТ-инфраструктуре и информационной безопасности, предварительной оценки зрелости, формирования клиентского заключения и внутреннего sales playbook.

Основной поток:

1. Заказчик заполняет анкету или продолжает ее из JSON-черновика.
2. Приложение нормализует ответы и проверяет фактические противоречия.
3. Gemini выполняет основной экспертный анализ, Groq используется как резервный провайдер.
4. Рекомендации сопоставляются с проверенным портфелем решений.
5. Формируется презентация и/или Excel-файл согласно настройкам администратора.
6. Внутренние материалы и диагностика отправляются в Telegram.
7. CRM-интеграция получает нормализованный лид после отдельного этапа активации.

### 2. Контуры и репозитории

| Контур | Назначение | Ссылка |
|---|---|---|
| Test | Разработка и приемка новых функций | https://testkh.streamlit.app |
| Админка Test | Управление настройками и интеграциями | https://testkh.streamlit.app/?admin=1 |
| GitHub Test | Исходный код предпроизводственного контура | https://github.com/irudoy1982/Test |
| GitHub KhalilAudit | Продакшен Khalil | https://github.com/irudoy1982/KhalilAudit |
| GitHub BTGAudit | Продакшен BTG | https://github.com/irudoy1982/BTGAudit |
| BTG Audit | Продакшен BTG | https://btgaudit.streamlit.app |

Все изменения сначала проверяются в Test. Перенос в KhalilAudit и BTGAudit выполняется только после ручной приемки.

### 3. Технологический стек

- Python 3.11 и Streamlit 1.41.1.
- pandas, openpyxl и XlsxWriter для обработки данных и Excel.
- requests для API-интеграций.
- Gemini API как основной ИИ-провайдер.
- Groq API как резервный ИИ-провайдер.
- Supabase как постоянное хранилище настроек, пользователей, файлов и журнала.
- Telegram Bot API для служебных уведомлений и доставки внутренних материалов.
- GitHub и Streamlit Community Cloud для поставки приложения.

### 4. Ключевые файлы

| Файл | Назначение |
|---|---|
| audit_app.py | Анкета, расчеты, ИИ-анализ, генерация результатов и Telegram |
| crm_admin.py | Админка, авторизация, настройки, пользователи и диагностика CRM |
| crm_store.py | Доступ к Supabase, защищенным настройкам и CRM |
| crm_delivery.py | Доставка результата в amoCRM, дедупликация, задача и вложения |
| crm_assets.py | Проверка логотипов, шаблонов презентаций и матрицы портфеля |
| db/*.sql | Последовательные миграции Supabase |
| static/ | Шаблоны презентаций, QR и брендовые материалы |
| vendor_matrix_detailed.xlsx | Резервная матрица решений и производителей |
| tools/ | Проверки, smoke-тесты и служебные генераторы |
| VERSIONING.md | Правила версий и roadmap |
| CHANGELOG_CUSTOMER.md | Только изменения, которые допустимо показывать заказчику |

### 5. Локальный запуск

1. Создать виртуальное окружение Python 3.11.
2. Установить зависимости: `pip install -r requirements.txt`.
3. Создать локальный `.streamlit/secrets.toml`.
4. Запустить `streamlit run audit_app.py`.
5. Открыть `http://localhost:8501`.

Не добавляйте `.streamlit/secrets.toml`, токены, пароли и service-role ключи в Git.

### 6. Secrets

| Ключ | Назначение |
|---|---|
| SUPABASE_URL | URL проекта Supabase |
| SUPABASE_SERVICE_ROLE_KEY | Серверный secret/service-role ключ Supabase |
| ADMIN_USERNAME | Начальный логин администратора |
| ADMIN_PASSWORD_HASH | PBKDF2-хеш начального пароля |
| TELEGRAM_TOKEN | Токен Telegram-бота |
| TELEGRAM_CHAT_ID | Чат для внутренних материалов и диагностики |
| GEMINI_API_KEY | Ключ основного ИИ-провайдера |
| GROQ_API_KEY | Ключ резервного ИИ-провайдера |

Настройка Streamlit Secrets: https://share.streamlit.io

### 7. Supabase

Создайте проект в https://supabase.com/dashboard и выполните миграции в SQL Editor строго по порядку:

1. `db/001_crm_admin.sql`
2. `db/002_admin_assets.sql`
3. `db/003_admin_users.sql`
4. `db/004_admin_password_recovery.sql`

RLS должен оставаться включенным. Service-role ключ используется только на сервере. После миграций добавьте `SUPABASE_URL` и `SUPABASE_SERVICE_ROLE_KEY` в Streamlit Secrets.

### 8. Генерация результатов

Клиентский результат выбирается в админке: презентация, Excel или оба файла. Sales playbook является внутренним материалом. ИИ получает технические ответы анкеты без контактных данных заказчика, кроме отраслевого контекста.

Порядок ИИ-провайдеров: Gemini, затем Groq. Если ни один провайдер не дал пригодный результат, клиентская выдача не должна маскировать проблему слабым материалом. Подробная диагностика остается внутренней.

Матрица портфеля должна содержать категории решений, производителей, дистрибьюторов и признак проверенного статуса. После замены файла используйте проверку в админке до публикации.

### 9. Админка и CRM

Админка Test: https://testkh.streamlit.app/?admin=1

Роли:

- administrator - пользователи, настройки, файлы, CRM и журнал;
- editor - рабочие настройки и файлы без управления администраторами;
- viewer - обзор и журнал без изменения конфигурации.

amoCRM в X3-dev.3 создает или переиспользует компанию и контакт, создает сделку и задачу, назначает ответственного и прикрепляет презентацию с Sales Playbook. Доставка защищена идемпотентным ключом и записывается в `crm_delivery_log`. Частичная ошибка вложений не блокирует клиентский результат. Bitrix24 подключается следующим адаптером после приемки amoCRM.

Документация amoCRM API: https://www.amocrm.ru/developers

### 10. Версионность

- `v12.33` - текущая стабильная производственная линия.
- `vX3-dev.N` - разработка в Test.
- `vX3-rcN` - кандидат после приемки.
- `vX3` - стабильный крупный выпуск.
- Число 13 записывается как X3.

Мелкие исправления увеличивают последнюю цифру. Новая функциональная стадия увеличивает номер dev-этапа.

### 11. Проверки перед релизом

1. `python -m py_compile audit_app.py crm_admin.py crm_store.py crm_delivery.py crm_assets.py`
2. `python tools/crm_delivery_smoke_test.py`
3. Запустить smoke-тесты черновика, анкеты, генерации и CRM-админки.
4. Проверить загрузку и повторное применение JSON.
5. Проверить презентацию в PowerPoint без восстановления файла.
6. Проверить Excel, Telegram, бренд и портфель.
7. Проверить мобильный экран и светлую тему.
8. Сверить номер версии и CHANGELOG_CUSTOMER.md.
9. Выполнить ручную приемку Test.
9. Только после подтверждения перенести изменения в два продакшен-репозитория.

### 12. Диагностика

- Streamlit не запускается: открыть Manage app, проверить логи, `requirements.txt`, `runtime.txt` и main file.
- Админка не работает: проверить URL `?admin=1`, Supabase Secrets и четыре миграции.
- Не приходит Telegram: проверить токен, chat ID, права бота и переключатели админки.
- ИИ не отвечает: проверить ключи, квоты Gemini и лимиты Groq.
- CRM отвечает 401: проверить домен и действующий access token.
- CRM отвечает 403 Admin access only: запросить права администратора amoCRM.
- Изменения не видны: проверить commit в нужном репозитории, версию приложения и deployment log.

### 13. Ссылки

- GitHub: https://github.com
- Streamlit Cloud: https://share.streamlit.io
- Supabase: https://supabase.com/dashboard
- Telegram Web: https://web.telegram.org
- BotFather: https://t.me/BotFather
- Gemini AI Studio: https://aistudio.google.com
- Gemini quotas: https://ai.dev/rate-limit
- Groq Console: https://console.groq.com
- amoCRM: https://www.amocrm.ru
- amoCRM API: https://www.amocrm.ru/developers
