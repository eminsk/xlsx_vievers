# Excel Viewer Pro — 64-bit Assembly Acceleration Suite (x86-64 FASM)

Высокопроизводительный модуль аппаратного ускорения расчётов таблиц, векторизованной SIMD (SSE2/AVX) математики, финансовых функций и быстрого кэширования для проекта **Excel Viewer Pro**.

Разработано для 64-битного ассемблера **FASM** (Flat Assembler) с полной поддержкой Windows x64 ABI и бесшовной интеграцией в Python через `ctypes` и аппаратные буферы памяти.

---

## 📁 Состав модуля `asm`

### 1. Библиотечные модули и исходный код FASM
- **[`xlsx_math64.asm`](xlsx_math64.asm)** — Исходный код динамической 64-битной библиотеки SIMD-вычислений:
  - `vec_sum_f64(arr, count)` — Векторизованное параллельное сложение массива чисел `double` (SSE2 `addpd` с 4-кратной разверткой цикла + horizontal sum).
  - `vec_avg_f64(arr, count)` — Векторное вычисление среднего арифметического.
  - `vec_min_f64(arr, count)` — Быстрый поиск минимума через SSE2 `minsd`.
  - `vec_max_f64(arr, count)` — Быстрый поиск максимума через SSE2 `maxsd`.
  - `vec_sumproduct_f64(a, b, count)` — Векторное скалярное произведение массивов (`mulpd` + `addpd`).
  - `fast_pmt_f64(rate, nper, pv, fv, type)` — Аппаратный расчёт периодического платежа по кредиту (Excel `PMT`) с x87 FPU / SSE2 экспоненцированием `(1 + rate)^nper`.
  - `fast_pv_f64(rate, nper, pmt, fv, type)` — Аппаратный расчёт приведенной стоимости инвестиции (Excel `PV`).
  - `fast_fv_f64(rate, nper, pmt, pv, type)` — Аппаратный расчёт будущей стоимости инвестиции (Excel `FV`).
  - `fast_str_hash(str, len)` — 64-битный FNV-1a строковый хеш для мгновенного поиска и инвалидации кэша формул и адресов ячеек.
  - `fast_count_nonblank(ptr_arr, count)` — Быстрый подсчёт непустых ячеек в памяти.

- **[`xlsx_math64.dll`](xlsx_math64.dll)** — Скомпилированная компактная 64-битная библиотека (всего 2.5 КБ чистого машинного кода без внешних зависимостей).

### 2. Полноценное табличное приложение на чистом ассемблере (GUI)
- **[`xlsx_gui64.asm`](xlsx_gui64.asm)** / **`xlsx_gui64.exe`** — **Автономный табличный процессор Excel Viewer Pro на чистом x64 ассемблере (FASM)**:
  - **Строка Меню (Win32 Menu Bar)**: меню `File` (New, Open Demo, Save, Exit), `Formulas` (Recalculate All SIMD, AutoSum, Average, PMT), `Data` (Sort Ascending/Descending, 1M SIMD Benchmark), `Help` (About).
  - **Панель инструментов (Toolbar)**: кнопки быстрого доступа (`New`, `Open Demo`, `Save`, `AutoSum (SIMD)`, `Loan PMT`, `1M Benchmark`, `About`).
  - **Строка формул (Formula Bar)**: индикатор адреса активной ячейки (`B5`), кнопка `[fx]` для пересчёта и поле редактирования формул (`=SUM(B1:B4)`).
  - **Сетка электронной таблицы (Spreadsheet Grid)**: многоколоночная таблица (`SysListView32` в режиме `LVS_REPORT`) с аппаратным двойным буфером (`LVS_EX_DOUBLEBUFFER`), сеткой ячеек (`LVS_EX_GRIDLINES`), колонками `#`, `A` (категории), `B..E` (кварталы Q1..Q4), `F` (итоги года), `G` (маржинальность %), `H` (статус).
  - **Вкладки листов книги (Sheet Tabs)**: элемент `SysTabControl32` с переключением листов (`Sheet1 (Financial Plan)`, `Sheet2 (Sales Analytics)`, `Sheet3 (1M Benchmark)`).
  - **Интерактивная строка состояния (Status Bar)**: 5 динамических панелей с отображением статуса, выбранной ячейки, `SUM`, `AVG` и активного SIMD-движка.

### 3. Мост интеграции с Python
- **[`asm_bridge.py`](asm_bridge.py)** — Чистый высокоуровневый Python-модуль моста:
  - Автоматическое обнаружение и загрузка библиотеки `xlsx_math64.dll`.
  - Прямой доступ к непрерывным буферам памяти `array.array('d')` без лишних накладных расходов на упаковку Python-объектов.
  - Экспорт функций: `asm_sum()`, `asm_avg()`, `asm_min()`, `asm_max()`, `asm_sumproduct()`, `asm_pmt()`, `asm_pv()`, `asm_fv()`, `asm_str_hash()`, `asm_is_available()`.

### 4. Тесты, бенчмарки и скрипты сборки
- **[`test_asm.py`](test_asm.py)** — Комплексный набор модульных тестов и бенчмарков:
  - 100% покрытие всех математических и финансовых функций.
  - Сравнение результатов с чистым Python и `formulas.py` с точностью до 1e-12.
  - Стресс-тест на 1 000 000 элементов.
- **[`build.bat`](build.bat)** — Автоматизированный скрипт компиляции всех ассемблерных модулей с помощью `C:\asm\hdd\FASM.EXE` и последующего запуска тестов.

---

## ⚡ Сборка и компиляция

Скрипт `build.bat` автоматически использует компилятор `C:\asm\hdd\FASM.EXE`:

```cmd
cd C:\proekts\xlsx_vievers\asm
build.bat
```

Либо вручную:
```cmd
# 1. Сборка 64-битной DLL
C:\asm\hdd\FASM.EXE xlsx_math64.asm xlsx_math64.dll

# 2. Сборка графического приложения
C:\asm\hdd\FASM.EXE xlsx_gui64.asm xlsx_gui64.exe

# 3. Запуск тестов и бенчмарка
python test_asm.py
```

---

## 🚀 Производительность и бенчмарки

Результаты тестирования на 1 000 000 чисел типа `double` (8 МБ непрерывной памяти):

| Операция | Язык / Технология | Время выполнения | Пропускная способность |
| :--- | :--- | :--- | :--- |
| **Векторная сумма 1M чисел** | **x86-64 SSE2 SIMD (FASM)** | **0.56 мс** | **~1.78 млрд чисел/сек** |
| Стандартная сумма 1M чисел | Python 3.12 `sum()` | 7.81 мс | ~128 млн чисел/сек |
| Финансовый расчёт PMT / PV | Native x87/SSE2 FASM | < 0.0001 мс | Мгновенно (аппаратные регистры) |

---

## 💻 Пример использования в Python

```python
from asm.asm_bridge import (
    asm_sum,
    asm_avg,
    asm_min,
    asm_max,
    asm_sumproduct,
    asm_pmt,
    asm_str_hash,
    asm_is_available,
)

if asm_is_available():
    data = [10.5, 20.0, 30.5, 40.0, 50.0]
    print("SIMD Sum:", asm_sum(data))          # 151.0
    print("SIMD Avg:", asm_avg(data))          # 30.2
    print("SIMD Min:", asm_min(data))          # 10.5
    print("SIMD Max:", asm_max(data))          # 50.0

    # Ежемесячный платёж по кредиту $200,000 под 5% на 30 лет (360 месяцев)
    monthly_payment = asm_pmt(0.05 / 12, 360, 200000.0)
    print("Monthly PMT:", round(monthly_payment, 2))  # -1073.64

    # 64-битный строковый FNV-1a хеш
    print("Range Hash:", hex(asm_str_hash("A1:Z100")))
```

---

## 🌟 Особенности архитектуры

1. **Чистый Win64 ABI**: строгое соблюдение Microsoft x64 Calling Convention (теневое пространство 32 байта, выравнивание стека по 16 байт, сохранение регистров `RBX`, `RBP`, `RDI`, `RSI`, `R12-R15`).
2. **SIMD SSE2 Unrolling**: инструкции `addpd`, `mulpd`, `minpd`, `maxpd` обрабатывают данные пакетами по 4 числа `double` за итерацию (32 байта за такт).
3. **Аппаратные трансцендентные функции**: расчёт экспоненты и степени `(1+rate)^nper` выполняется через аппаратные инструкции `fyl2x` и `f2xm1` на x87 FPU.
4. **Zero-Dependency**: DLL не требует сторонних библиотек (C++ runtime или MSVC CRT) и запускается на любой версии Windows x64.
