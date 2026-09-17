import os
import sys
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx_presentation_builder import EditablePPTXBuilder


def create_presentation(output_path):
    builder = EditablePPTXBuilder(theme="dark")

    COLOR_CARD = builder.COLOR_CARD
    COLOR_TEXT_MAIN = builder.COLOR_TEXT_MAIN
    COLOR_TEXT_MUTED = builder.COLOR_TEXT_MUTED
    COLOR_BLUE = builder.COLOR_PRIMARY

    # -------------------------------------------------------------
    # SLIDE 1: Title Slide
    # -------------------------------------------------------------
    slide1 = builder.add_blank_slide()

    # Header tag
    tag = slide1.shapes.add_textbox(Inches(1.0), Inches(1.2), Inches(11.3), Inches(0.5))
    p = tag.text_frame.paragraphs[0]
    p.text = "ФГАОУ ВО МГТУ «СТАНКИН» • 2026"
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = COLOR_BLUE

    # Title
    tbox = slide1.shapes.add_textbox(Inches(1.0), Inches(1.8), Inches(11.3), Inches(1.8))
    tf = tbox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Проблема синхронизации данных между облачными и локальными хранилищами"
    p.font.size = Pt(34)
    p.font.bold = True
    p.font.color.rgb = COLOR_TEXT_MAIN

    # Subtitle
    sub = slide1.shapes.add_textbox(Inches(1.0), Inches(3.7), Inches(11.3), Inches(1.0))
    p = sub.text_frame.paragraphs[0]
    p.text = "Анализ когерентности, версионных коллизий, теорем CAP/PACELC и алгоритмов блочной дедупликации в цифровом производстве"
    p.font.size = Pt(16)
    p.font.color.rgb = COLOR_TEXT_MUTED

    # 3 info cards
    builder.add_card(slide1, Inches(1.0), Inches(4.8), Inches(3.5), Inches(1.8), "Дисциплина", ["Основы технологий цифрового производства"], icon="⚡")
    builder.add_card(slide1, Inches(4.9), Inches(4.8), Inches(3.5), Inches(1.8), "Фокус работы", ["Edge & Cloud Hybrid Systems"], icon="⚙️")
    builder.add_card(slide1, Inches(8.8), Inches(4.8), Inches(3.5), Inches(1.8), "Формат", ["Академический доклад по ГОСТу"], icon="🎓")
    builder.add_footer(slide1, 1)

    # -------------------------------------------------------------
    # SLIDE 2: Введение & Гибридная Архитектура
    # -------------------------------------------------------------
    slide2 = builder.add_blank_slide()
    builder.add_header(slide2, "01. Архитектура", "Предпосылки гибридного хранения данных", "Разделение информационных потоков: Локальный контур vs Публичное облако")

    builder.add_card(slide2, Inches(0.8), Inches(2.2), Inches(5.6), Inches(4.4),
                "Локальный контур (Edge / NAS / ПЛК)", [
                    "Управление оборудованием: ПЛК и IIoT требуют отклика < 10 мс без сетевого джиттера.",
                    "Автономия производства: Обрыв внешней связи не должен останавливать работу цеха.",
                    "Безопасность КИИ: Рецептуры и конструкторские чертежи защищены коммерческой тайной.",
                    "Первичная обработка: Фильтрация и сжатие сырых сигналов вибраций и температур."
                ], icon="🏭")

    builder.add_card(slide2, Inches(6.8), Inches(2.2), Inches(5.6), Inches(4.4),
                "Облачный контур (Cloud S3 / Analytics)", [
                    "Предиктивная аналитика: Обучение ML-моделей на долгосрочных архивах телеметрии.",
                    "Глобальный обмен: Централизованная синхронизация НСИ и документации между филиалами.",
                    "Неограниченные ресурсы: Масштабируемые мощности объектных хранилищ (Amazon S3, Azure).",
                    "Резервное копирование: Долгосрочное «холодное» хранение снимков данных."
                ], icon="☁️")
    builder.add_footer(slide2, 2)

    # -------------------------------------------------------------
    # SLIDE 3: Теоремы CAP и PACELC
    # -------------------------------------------------------------
    slide3 = builder.add_blank_slide()
    builder.add_header(slide3, "02. Теория систем", "Ограничения теорем CAP и PACELC", "Фундаментальный выбор между задержкой (Latency) и согласованностью (Consistency)")

    builder.add_card(slide3, Inches(0.8), Inches(2.2), Inches(5.6), Inches(4.4),
                "Теорема CAP в гибридных средах", [
                    "В распределенной системе при наличии сетевых разрывов (P) выбор сводится к CP или AP.",
                    "Модель CP (Consistency/Partition): При обрыве связи узел БЛОКИРУЕТ запись. Неприемлемо для ЧПУ!",
                    "Модель AP (Availability/Partition): Узел продолжает запись автономно. Данные расходятся → неизбежно последующее слияние."
                ], icon="⚖️")

    builder.add_card(slide3, Inches(6.8), Inches(2.2), Inches(5.6), Inches(4.4),
                "Теорема PACELC & Eventual Consistency", [
                    "В штатном режиме (Else) выбирается компромисс между Latency (L) и Consistency (C).",
                    "Strong Consistency: Задержка 100+ мс на каждый RTT запрос к облаку разрушает производительность I/O.",
                    "Eventual Consistency: Мгновенная локальная запись + асинхронный сброс в облако.",
                    "Inconsistency Window: Окно несогласованности, где возникают версионные конфликты."
                ], icon="⏱️")
    builder.add_footer(slide3, 3)

    # -------------------------------------------------------------
    # SLIDE 4: Ключевые коллизии и конфликты
    # -------------------------------------------------------------
    slide4 = builder.add_blank_slide()
    builder.add_header(slide4, "03. Анализ коллизий", "Ключевые проблемы и аномалии синхронизации", "Аномалии параллельной модификации, дрейф часов и избыточность трафика")

    builder.add_card(slide4, Inches(0.8), Inches(2.2), Inches(2.7), Inches(4.4), "Lost Update", [
        "Параллельные изменения локального цеха и облака.",
        "Стратегия Last-Write-Wins безвозвратно затирает правки."
    ], icon="💥")

    builder.add_card(slide4, Inches(3.8), Inches(2.2), Inches(2.7), Inches(4.4), "Clock Drift", [
        "Дрейф физических часов таймеров серверов.",
        "Метки времени NTP искажают хронологию операций."
    ], icon="🕰️")

    builder.add_card(slide4, Inches(6.8), Inches(2.2), Inches(2.7), Inches(4.4), "Whole-File Sync", [
        "Пересылка 10 ГБ файла из-за правки 2 байт.",
        "Перегрузка интернет-канала и раздувание egress-трафика."
    ], icon="📦")

    builder.add_card(slide4, Inches(9.8), Inches(2.2), Inches(2.7), Inches(4.4), "POSIX vs S3", [
        "Несовместимость метаданных POSIX ACL и S3 Object API.",
        "Потеря расширенных атрибутов xattr и жестких ссылок."
    ], icon="🗂️")
    builder.add_footer(slide4, 4)

    # -------------------------------------------------------------
    # SLIDE 5: Логическое время и CRDT
    # -------------------------------------------------------------
    slide5 = builder.add_blank_slide()
    builder.add_header(slide5, "04. Алгоритмы арбитража", "Векторные часы и структуры CRDT", "Детерминированный арбитраж конфликтов без использования физического времени")

    builder.add_card(slide5, Inches(0.8), Inches(2.2), Inches(5.6), Inches(4.4),
                "Векторные часы (Version Vectors)", [
                    "Каждый узел ведет вектор счетчиков V = [c1, c2, ... cn].",
                    "Точно определяет отношение причинно-следственной связи (Causality).",
                    "V2 >= V1 → Прямой потомок (безопасное обновление).",
                    "V2 < V1 → Устаревший пакет данных (игнорирование).",
                    "Разнонаправленные векторы → Фиксация параллельного конфликта."
                ], icon="🔢")

    builder.add_card(slide5, Inches(6.8), Inches(2.2), Inches(5.6), Inches(4.4),
                "Бесконфликтные типы данных (CRDT)", [
                    "Гарантируют строгую сходимость в конечном счете (Strong Eventual Consistency).",
                    "1. Коммутативность: A ⊔ B = B ⊔ A (порядок пакетов не имеет значения).",
                    "2. Ассоциативность: (A ⊔ B) ⊔ C = A ⊔ (B ⊔ C).",
                    "3. Идемпотентность: A ⊔ A = A (повторные пакеты не портят состояние).",
                    "Применение: Распределенные счетчики и текстовые документы."
                ], icon="🧬")
    builder.add_footer(slide5, 5)

    # -------------------------------------------------------------
    # SLIDE 6: Контентно-зависимый чанкинг (FastCDC)
    # -------------------------------------------------------------
    slide6 = builder.add_blank_slide()
    builder.add_header(slide6, "05. Оптимизация трафика", "Контентно-зависимое разбиение (FastCDC)", "Блочная дельта-синхронизация и устранение проблемы сдвига границ")

    builder.add_card(slide6, Inches(0.8), Inches(2.2), Inches(5.6), Inches(4.4),
                "Проблема фиксированного чанкинга", [
                    "При разбиении на фиксированные блоки (по 4 МБ) вставка 1 байта в начало файла сдвигает все границы.",
                    "Контрольные суммы абсолютно всех последующих блоков меняются.",
                    "Система вынуждена повторно передавать 100% объема файла по сети.",
                    "Нерациональный расход полосы пропускания и ресурсов хранилища."
                ], icon="✂️")

    builder.add_card(slide6, Inches(6.8), Inches(2.2), Inches(5.6), Inches(4.4),
                "Алгоритм FastCDC / Rabin Fingerprinting", [
                    "Точки разреза фиксируются структурой содержимого с помощью скользящего окна (48-64 байт).",
                    "Скользящее хеширование за O(1) при сдвиге на 1 байт.",
                    "Граница блока фиксируется, когда младшие k бит хеша равны 0.",
                    "Вставка данных меняет границы только 1 текущего блока! Все остальные блоки сохраняют свои SHA-256 ID."
                ], icon="⚡")
    builder.add_footer(slide6, 6)

    # -------------------------------------------------------------
    # SLIDE 7: Деревья Меркла (Merkle Trees)
    # -------------------------------------------------------------
    slide7 = builder.add_blank_slide()
    builder.add_header(slide7, "06. Быстрая сверка", "Иерархическая сверка через Деревья Меркла", "Локализация расхождений данных за O(log N) сетевых запросов")

    builder.add_card(slide7, Inches(0.8), Inches(2.2), Inches(7.0), Inches(4.4),
                "Принцип работы Дерева Хешей (Merkle Tree)", [
                    "Листья дерева содержат криптографические хеши (SHA-256) отдельных блоков файла.",
                    "Родительские узлы вычисляются путем хеширования пары дочерних хешей.",
                    "1. Сначала локальный агент и облако сверяют только Root Hash (1 сетевой запрос).",
                    "2. Если Root Hash совпал — массивы данных 100% идентичны!",
                    "3. При расхождении система спускается только по ветвям с изменившимися хешами."
                ], icon="🌳")

    builder.add_card(slide7, Inches(8.2), Inches(2.2), Inches(4.2), Inches(4.4),
                "Выигрыш в производительности", [
                    "Сложность поиска: O(log N) вместо полного сканирования O(N).",
                    "Экономия трафика: Сокращение служебных запросов более чем на 95%.",
                    "Безопасность: Криптографическая гарантированная целостность каждого блока."
                ], icon="📊")
    builder.add_footer(slide7, 7)

    # -------------------------------------------------------------
    # SLIDE 8: CDC & Event-Driven Architecture
    # -------------------------------------------------------------
    slide8 = builder.add_blank_slide()
    builder.add_header(slide8, "07. CDC & Event-Driven", "Захват изменений данных (CDC & Kafka)", "Переход от периодического опроса (Polling) к реагированию на события ядра ОС")

    builder.add_card(slide8, Inches(0.8), Inches(2.2), Inches(3.6), Inches(4.4), "Ядро ОС (Inotify / USN)", [
        "Linux VFS: inotify и fanotify перехватывают вызовы modify, create, delete.",
        "Windows NTFS: ReadDirectoryChangesW и журнал USN."
    ], icon="🐧")

    builder.add_card(slide8, Inches(4.8), Inches(2.2), Inches(3.6), Inches(4.4), "СУБД WAL / Debezium", [
        "Прямое чтение транзакционного журнала (WAL в Postgres, binlog в MySQL).",
        "Без тяжелых SELECT запросов!"
    ], icon="🗄️")

    builder.add_card(slide8, Inches(8.8), Inches(2.2), Inches(3.6), Inches(4.4), "Apache Kafka / Bus", [
        "Потоковая трансляция событий с семантикой At-Least-Once.",
        "Фиксация Offsets позволяет легко восстановить чтение после сбоя."
    ], icon="📨")
    builder.add_footer(slide8, 8)

    # -------------------------------------------------------------
    # SLIDE 9: Cloud Storage Gateways
    # -------------------------------------------------------------
    slide9 = builder.add_blank_slide()
    builder.add_header(slide9, "08. Шлюзы хранения", "Cloud Storage Gateways & Automated Tiering", "Локальный кэш записи и прозрачное разрежение файлов (Stub Files)")

    builder.add_card(slide9, Inches(0.8), Inches(2.2), Inches(5.6), Inches(4.4),
                "Трансляция протоколов & Write-Back Cache", [
                    "Предоставляет цеху стандартные интерфейсы SMB / NFS / iSCSI.",
                    "С обратной стороны транслирует операции в REST API S3.",
                    "Write-Back Cache (NVMe): Запись подтверждается локально за < 1 мс.",
                    "Асинхронный воркер сжимает и сбрасывает данные в облако."
                ], icon="🔀")

    builder.add_card(slide9, Inches(6.8), Inches(2.2), Inches(5.6), Inches(4.4),
                "Разряжение файлов (Cloud Tiering)", [
                    "Редко используемые «холодные» файлы вытесняются в облачное хранилище.",
                    "Локально остается Stub File (Заглушка), сохраняющий имя и POSIX/NTFS ACL.",
                    "File Hydration: При обращении к заглушке драйвер прозрачно подгружает требуемый Byte-Range из S3."
                ], icon="❄️")
    builder.add_footer(slide9, 9)

    # -------------------------------------------------------------
    # SLIDE 10: Сравнительный анализ систем (Table)
    # -------------------------------------------------------------
    slide10 = builder.add_blank_slide()
    builder.add_header(slide10, "09. Сравнение решений", "Сравнительный анализ промышленных систем", "Сравнение AWS Storage Gateway, Azure File Sync, Nextcloud, Seafile и MinIO")

    # Create Table
    rows = 6
    cols = 5
    left = Inches(0.8)
    top = Inches(2.2)
    width = Inches(11.7)
    height = Inches(4.4)

    table_shape = slide10.shapes.add_table(rows, cols, left, top, width, height)
    table = table_shape.table

    headers = ["Платформа", "Модель синхронизации", "Блочная дельта-синхр.", "Арбитраж конфликтов", "Права доступа"]
    for i, h in enumerate(headers):
        cell = table.cell(0, i)
        cell.text = h
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(30, 41, 64)
        for p in cell.text_frame.paragraphs:
            p.font.bold = True
            p.font.size = Pt(12)
            p.font.color.rgb = COLOR_TEXT_MAIN

    data = [
        ["Azure File Sync", "Hub-and-Spoke", "Поблочно (4 МБ)", "Конфликтные копии", "NTFS ACL (Полная)"],
        ["AWS Storage Gateway", "Файловый / Блочный", "Да (EBS / S3 Tiering)", "Версионирование S3", "POSIX / SMB ACL"],
        ["Nextcloud", "Двунаправленная", "Нет (Файл целиком)", "Версионирование в БД", "Базовые WebDAV"],
        ["Seafile", "Деревья коммитов", "Да (FastCDC чанкинг)", "Ветвление репозитория", "Собственные ACL"],
        ["MinIO", "S3 Replication", "Объектная репликация", "S3 Versioning", "IAM Политики"]
    ]

    for row_idx, row_data in enumerate(data, start=1):
        for col_idx, text in enumerate(row_data):
            cell = table.cell(row_idx, col_idx)
            cell.text = text
            cell.fill.solid()
            cell.fill.fore_color.rgb = COLOR_CARD
            for p in cell.text_frame.paragraphs:
                p.font.size = Pt(11)
                p.font.color.rgb = COLOR_TEXT_MUTED
                if col_idx == 0:
                    p.font.bold = True
                    p.font.color.rgb = COLOR_TEXT_MAIN

    builder.add_footer(slide10, 10)

    # -------------------------------------------------------------
    # SLIDE 11: Практические кейсы & Рекомендации
    # -------------------------------------------------------------
    slide11 = builder.add_blank_slide()
    builder.add_header(slide11, "10. Рекомендации", "Практические кейсы внедрения в производстве", "Оптимальный выбор архитектуры под разные типы задач предприятия")

    builder.add_card(slide11, Inches(0.8), Inches(2.2), Inches(3.6), Inches(4.4), "САПР & 3D-Модели", [
        "Решение: Seafile / FastCDC.",
        "Передаются только дельта-блоки при изменении параметров чертежей.",
        "Защищает каналы связи от многогигабайтных повторных загрузок."
    ], icon="📐")

    builder.add_card(slide11, Inches(4.8), Inches(2.2), Inches(3.6), Inches(4.4), "MES & Станки ЧПУ", [
        "Решение: Гибридный шлюз (SMB/NFS).",
        "Станки работают с локальной папкой без установки агентов.",
        "Полная автономия цеха при разрывах интернета."
    ], icon="⚙️")

    builder.add_card(slide11, Inches(8.8), Inches(2.2), Inches(3.6), Inches(4.4), "IIoT & Телеметрия", [
        "Решение: MinIO + Kafka Bus.",
        "Локальное накопление рядов данных с асинхронной репликацией бакетов.",
        "Обучение ML-моделей в облаке."
    ], icon="📡")
    builder.add_footer(slide11, 11)

    # -------------------------------------------------------------
    # SLIDE 12: Заключение
    # -------------------------------------------------------------
    slide12 = builder.add_blank_slide()
    builder.add_header(slide12, "11. Выводы", "Заключение и ключевые результаты", "Научно-практические выводы исследования синхронизации данных")

    builder.add_card(slide12, Inches(0.8), Inches(2.2), Inches(6.0), Inches(4.4),
                "Основные результаты исследования", [
                    "Eventual Consistency — единственная практичная модель для гибридных производств.",
                    "Логическое время (CRDT) полностью устраняет ошибки дрейфа физических часов NTP.",
                    "FastCDC & Merkle Trees дают экономию сетевого трафика более 90%.",
                    "Гибридные шлюзы (Storage Gateways) обеспечивают автономию цеха при авариях связи."
                ], icon="✅")

    # Final box
    final_card = slide12.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.2), Inches(2.2), Inches(5.3), Inches(4.4))
    final_card.fill.solid()
    final_card.fill.fore_color.rgb = RGBColor(30, 45, 75)
    final_card.line.color.rgb = COLOR_BLUE
    final_card.line.width = Pt(2)

    tf = final_card.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Спасибо за внимание!"
    p.font.size = Pt(28)
    p.font.bold = True
    p.font.color.rgb = COLOR_TEXT_MAIN
    p.alignment = PP_ALIGN.CENTER
    p.space_after = Pt(20)

    p2 = tf.add_paragraph()
    p2.text = "Готов ответить на ваши вопросы\nпо теме доклада"
    p2.font.size = Pt(16)
    p2.font.color.rgb = COLOR_TEXT_MUTED
    p2.alignment = PP_ALIGN.CENTER
    p2.space_after = Pt(30)

    p3 = tf.add_paragraph()
    p3.text = "МГТУ «СТАНКИН» • 2026"
    p3.font.size = Pt(14)
    p3.font.bold = True
    p3.font.color.rgb = COLOR_BLUE
    p3.alignment = PP_ALIGN.CENTER

    builder.add_footer(slide12, 12)

    # Save presentation
    builder.save(output_path)


if __name__ == "__main__":
    output_file = sys.argv[1] if len(sys.argv) > 1 else "03_Готовые_доклады/Презентация_Проблема_синхронизации_данных.pptx"
    os.makedirs(os.path.dirname(output_file), exist_ok=True)
    create_presentation(output_file)
