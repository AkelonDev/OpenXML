<h1>Паспорт шаблона разработки «OpenXML»</h1>
<h2>Возможности решения</h2>
Решение представляет набор функциональности по заполнению основных свойств документов формата ".docx". 
Решение позволяет подставлять необходимые значения в элементы управления содержимым «Обычный текст», «Форматированный текст» и «Рисунок». 
<h3>А также:</h3>
1.	Создавать, заполнять и вставлять новые таблицы на место «Закладки».</br>
2.	Вставлять новые значения/картинки в ячейки таблиц.</br>
3.	Заменять текст в теле документа.</br>
4.	Генерировать штрих/qr коды для дальнейшей подстановки в документ.</br>
5.	Добавлять подложку и изменять её значение.</br>
<h3>Апробация на проектах</h3>
Использованные в решении подходы применялись на проектах:</br>
•	Проект 1 — Р-Фарм- Внедрение DirectumRX;</br>
•	Проект 2 — Микроген - Внедрение DirectumRX;</br>
•	Проект 3 — Мерц Фарм - Внедрение DirectumRX;</br>
•	И многих других.</br>
<h3>Состав решения</h3>
1.	Модуль OpenXML.</br>
2.	Изолированная область (AkelonOpenXMLWrapper).</br>
3.	Шаблон для демонстрации решения (создаётся при инициализации).</br>
4.	Используемые внешние библиотеки: DocumentFormat.OpenXml.dll – версия 3.0.1, DocumentFormat.OpenXml.Framework.dll – версия 3.0.1, System.IO.Packaging.dll – версии 8.0.0.</br>
5. Обожка модуля с действием для демонстрации решения.
<h3>Варианты расширения функциональности на проектах</h3>
На проекте при необходимости можно добавлять перегрузки для существующих методов. Например, для добавления управления настройкой курсива для значений заполняемых свойств. 
А так же добавлять новые методы в изолированную область для расширения функционала.
<h3>Архитектурно неочевидные моменты</h3>
1.	Чтобы сделать подложку иного стиля:</br>
 <ul>
  a)	Создать новый документ и в нём настроить вид подложки.</br>
  b)	Установить OpenXML SDK 2.5 и открыть с помощью него документ с нужной подложкой.</br>
  c)	Найти фрагмент кода с генерацией подходящей подложки. Проще всего искать по тексту подложки.
 </ul>
2. Шаблон для демонстарции решения вместе с версией создаётся при инициализации.
