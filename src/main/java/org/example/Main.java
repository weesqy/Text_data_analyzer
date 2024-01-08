package org.example;

import org.dhatim.fastexcel.Workbook;
import org.dhatim.fastexcel.Worksheet;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.logging.log4j.Logger;

import java.io.File;
import java.io.FileOutputStream;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.time.LocalDateTime;
import java.util.*;

/**
 * Метод для запуска приложений
 */
public class Main {

    // Создаём объект логгера
    final static Logger logger = LogManager.getLogger(String.valueOf(Main.class));

    public static void main(String[] args) {
        logger.info("Начало исполнения программы:" + LocalDateTime.now());

        // Получим разрешения файлов для проверки, которые вводит юзер
        var inputExtensions = GetExtensionsFromConsole();

        // Получим имена файлов их папки filesToCheck. Берём только файлы, у которых расширение введённое юзером
        var fileNamesToCheck = GetFileNamesToCheckByInputExtensions(inputExtensions);

        // Получим имена словарей, лежащих в папке dictionaries
        var dictionaryFileNames = GetDictionaryFileNames();

        // Получим словарь-java формата: ключ - тема словаря, значение - массив слов в словаре
        var checkingDictionaries = GetCheckingDictionaries(dictionaryFileNames);

        // Запускаем цикл по именам файлов для проверки
        for (String fileName : fileNamesToCheck) {
            // Оборачиваем код в конструкцию try catch, чтобы обработать возможную ошибку
            try {
                // Считаем файл для проверки в переменную типа string
                String checkContent = new String(Files.readAllBytes(Paths.get(".\\filesToCheck\\" + fileName)));

                // Создадим файл Excel на запись
                var file = new File(fileName.substring(0, fileName.lastIndexOf('.')) + ".xlsx");

                // Откроем поток созданного файла на запись. Оборачиваем код в конструкцию try catch, чтобы обработать возможную ошибку
                try (var fos = new FileOutputStream(file)) {

                    // Создадим счётчик (словарь-java) формата: ключ - имя словаря, значение - количество вхождений слов словаря в файл для проверки
                    var fullCounter = new HashMap<String, Integer>();

                    // Создадим рабочее пространство Excel с записью в потом созданного файла
                    var wb = new Workbook(fos, "Application", "1.0");

                    // Создадим лист в Excel с названием "Result"
                    Worksheet sheet = wb.newWorksheet("Result");

                    // Объединим клетки: столбцы с 4 по 10, строки с 1 по 1 (то есть строка и так одна)
                    sheet.range(1, 4, 1, 10).merge();

                    // Вставим в объединённую клетку: столбцы с 4 по 10, строки с 1 по 1 текст "Исходный текст"
                    sheet.value(1, 4, "Исходный текст");

                    // Объединим клетки: столбцы с 4 по 10, строки с 2 по 8
                    sheet.range(2, 4, 8, 10).merge();

                    // Вставим в объединённую клетку: столбцы с 4 по 10, строки с 2 по 8 текст из файла для проверки
                    sheet.value(2, 4, checkContent);

                    // Зададим стиль клеток: столбцы с 4 по 10, строки с 1 по 8:
                    // Задать границу клеток, задать горизонтальную и вертикальную ориентации клеток
                    sheet.range(1, 4, 8, 10).style().borderStyle("thin")
                            .horizontalAlignment("center").verticalAlignment("center").wrapText(true).set();

                    // Преобразуем текст из файла для проверки в нижний регистр для дальнейшей проверки без учёта регистра
                    checkContent = checkContent.toLowerCase();

                    // Запускаем цикл по словарь-java формата ключ - тема словаря, значение - массив слов в словаре
                    for (Map.Entry<String, List<String>> entry : checkingDictionaries.entrySet()) {

                        // Создадим счётчик количества вхождений слов конкретного словаря в конкретный файл для проверки
                        var dictCount = 0;

                        // Создадим лист в Excel с названием как у темы словаря
                        Worksheet sheetC = wb.newWorksheet(entry.getKey());

                        // Зададим значение ячейки (1, 1) (не с 0,0, чтобы был отступ сверху и снизу в одну клеточку). Запишем туда "Слово"
                        sheetC.value(1, 1, "Слово");

                        // Зададим ширину столбца 1
                        sheetC.width(1, 22);

                        // Зададим значение ячейки (1, 2). Запишем туда "Количество вхождений"
                        sheetC.value(1, 2, "Количество вхождений");

                        // Зададим ширину столбца 2
                        sheetC.width(2, 22);

                        // Зададим счётчик строки Excel (начинает с 2 строки, тк выше уже создали 1)
                        var row = 2;

                        // Запускаем цикл по массиву слов в словаре
                        for (String word : entry.getValue()) {

                            if(word.isEmpty()) {
                                continue;
                            }

                            // Заменим в тексте файла для проверки все встречающиеся слова word на пустую строку
                            var replacedText = checkContent.replace(word, "");

                            // Посчитаем разницу длины текста исходного файла со словом word и длины текста исходного файла без слова word,
                            // затем разделим эту разницу на длину слова word. Таким образом получим количество вхождений слова в исходном тексте
                            var count = (checkContent.length() - replacedText.length()) / word.length();

                            // Если вхождений слова в текст для проверки нет, то не записываем нулевое количество вхождений в Excel, а идём дальше
                            if (count == 0)
                                continue;

                            // Зададим значение ячейки (row, 1). Запишем туда слово word
                            sheetC.value(row, 1, word);

                            // Зададим значение ячейки (row, 2). Запишем туда количество вхождений слова word в конкретный словарь
                            sheetC.value(row, 2, count);

                            // Увеличим счётчик количества вхождений слов конкретного словаря в конкретный файл для проверки
                            dictCount += count;

                            // Увеличим номер строки для записи в Excel
                            row++;
                        }

                        // Зададим значение ячейки (row, 1). Запишем туда "Всего:"
                        sheetC.value(row, 1, "Всего:");

                        // Зададим значение ячейки (row, 2). Запишем туда количество вхождений всех слов конкретного словаря в конкретный файл для проверки
                        sheetC.value(row, 2, dictCount);

                        // Зададим стиль клеток: столбцы с 2 по 2 (то есть столбец 2), строки с 2 по row:
                        // Задать горизонтальную и вертикальную ориентации клеток
                        sheetC.range(2, 2, row, 2).style()
                                .horizontalAlignment("center").verticalAlignment("center").set();

                        // Зададим стиль клеток: столбцы с 1 по 2, строки с 1 по row:
                        // Задать границу
                        sheetC.range(1, 1, row, 2).style().borderStyle("thin").set();

                        // Добавим в счётчик словарь-java элемент: ключ - наименование словаря, значение - количество вхождений слов словаря в конкретный файл для проверки
                        fullCounter.put(entry.getKey(), dictCount);
                    }

                    // Зададим значение ячейки (1, 1). Запишем туда "Вычисленная тема:"
                    sheet.value(1, 1, "Вычисленная тема:");
                    // Зададим ширину столбца 1
                    sheet.width(1, 20);

                    // Получим название файла словаря, у которого больше всего вхождений слов в конкретный файл для проверки
                    var calculateDict = Collections.max(fullCounter.entrySet(), Map.Entry.comparingByValue()).getKey();
                    // Уберём расширение названия файла словаря и постфикс (уберём _dictionary у "Имя_файла_dictionary.txt")
                    var calculatedTopic = calculateDict.substring(0, calculateDict.lastIndexOf("_"));
                    // Зададим значение ячейки (1, 2). Запишем туда название файла словаря, у которого больше всего вхождений слов в конкретный файл для проверки
                    sheet.value(1, 2, calculatedTopic);

                    // Зададим ширину столбца 2
                    sheet.width(2, 20);

                    // Зададим стиль клеток: столбцы с 1 по 2, строки с 1 по row:
                    // Задать границу
                    sheet.range(1, 1, 1, 2).style().borderStyle("thin").set();

                    // Зададим значение ячейки (3, 1). Запишем туда "Словарь"
                    sheet.value(3, 1, "Словарь");
                    // Зададим значение ячейки (3, 2). Запишем туда "Вхождения слов"
                    sheet.value(3, 2, "Вхождения слов");

                    // Зададим счётчик строки Excel (начинает с 4 строки, тк выше уже создали первые 3 строки)
                    var row = 4;
                    // Запускаем цикл по счётчику (словарю-java) формата: ключ - имя словаря, значение - количество вхождений слов словаря в файл для проверки
                    for (Map.Entry<String, Integer> entry : fullCounter.entrySet()) {
                        // Получим имя файла словаря
                        var dict = entry.getKey();
                        // Зададим значение ячейки (row, 1). Запишем туда название словаря без конца "_dictionary.txt"
                        sheet.value(row, 1, dict.substring(0, dict.lastIndexOf("_")));
                        // Зададим значение ячейки (row, 2). Запишем туда количество вхождений слов словаря в файл для проверки
                        sheet.value(row, 2, entry.getValue());
                        // Увеличим номер строки для записи в Excel
                        row++;
                    }

                    // Зададим стиль клеток: столбцы с 2 по 2 (то есть столбец 2), строки с 4 по row-1:
                    // Задать границу
                    sheet.range(4, 2, row - 1, 2).style()
                            .horizontalAlignment("center").verticalAlignment("center").set();

                    // Зададим стиль клеток: столбцы с 1 по 2, строки с 3 по row-1:
                    // Задать горизонтальную и вертикальную ориентации клеток
                    sheet.range(3, 1, row - 1, 2).style().borderStyle("thin").set();

                    // Закроем рабочее пространство Excel
                    wb.finish();
                }
                // Если не удалось считать данные из конкретного файла для проверки - бросаем ошибку
            } catch (Exception e) {
                var errorStr = "Не удалось считать файл с именем \"" + fileName + '\"';
                logger.error(errorStr);
                throw new IllegalArgumentException(errorStr);
            }
            logger.info("Конец исполнения программы:" + LocalDateTime.now());
        }
    }

    /**
     * Возвращает словарь-java формата: ключ - тема словаря, значение - массив слов в словаре
     * @param dictionaryFileNames Имена файлов словарей для проверки
     * @return Java-словарь содержащий информацию о словарях для проверки
     */
    private static Map<String, List<String>> GetCheckingDictionaries(List<String> dictionaryFileNames) {
        // Создадим словарь-java формата: ключ - тема словаря, значение - массив слов в словаре
        var map = new HashMap<String, List<String>>();
        // Запускаем цикл по именованиям словарей файлов
        for (String fileName : dictionaryFileNames) {
            // Оборачиваем код в конструкцию try catch, чтобы обработать возможную ошибку
            try {
                // Считываем содержимое словаря в переменную типа string
                String checkContent = new String(Files.readAllBytes(Paths.get(".\\dictionaries\\" + fileName)));
                // Преобразуем считанный контент в нижний регистр для дальнейшей проверки вне зависимости от регистра
                checkContent = checkContent.toLowerCase();
                // Разделим строку по переходам в новую строку и запишем в лист, затем добавим всё это в словарь
                var list = List.of(checkContent.split("\n"));
                var distinctList = new ArrayList<>(new HashSet<>(list));
                map.put(fileName, distinctList);
                // Если не удалось считать данные из конкретного словаря - бросаем ошибку
            } catch (Exception e) {
                var errorStr = "Не удалось считать словарь с именем \"" + fileName + '\"';
                logger.error(errorStr);
                throw new IllegalArgumentException(errorStr);
            }
        }

        // Возвращаем словарь из функции
        return map;
    }

    /**
     *  Возвращает имена файлов словарей
     * @return Список имён файлов
     */
    private static List<String> GetDictionaryFileNames() {

        // Зададим требуемое окончание словарей
        String dictPostFix = "_dictionary.txt";

        // Зададим папку, откуда будем считывать словари
        File folder = new File(".\\dictionaries");

        // Считаем все файлы папки
        var files = folder.listFiles();

        // Если папки нет, то выдадим соответствующую ошибку
        if (files == null){
            var errorStr = "\nОтсутствует папка \"..\\dictionaries\"";
            logger.error(errorStr);
            throw new NullPointerException(errorStr);
        }

        // Если в папке нет файлов, то выдадим соответствующую ошибку
        if (files.length == 0){
            var errorStr = "\nОтсутствуют файлы в папке \"..\\dictionaries\"";
            logger.error(errorStr);
            throw new IllegalArgumentException(errorStr);
        }

        // Зададим лист, где будем записывать названия файлов словарей, которые будут использованы в проверке
        List<String> dictionaryNames = new ArrayList<>();
        // Зададим цикл по файлам в папке
        for (File file : files) {
            // Получим имя файла
            var fileName = file.getName();
            // Если имя словаря оканчивается на "_dictionary.txt", то добавим имя файла в лист
            if (fileName.endsWith(dictPostFix))
                dictionaryNames.add(fileName);
        }

        // Если были найдены словари, выведем сообщение в консоль
        if (!dictionaryNames.isEmpty()) {
            System.out.println("Найдены словари:");
            for (String fileName : dictionaryNames) {
                System.out.println(fileName);
            }
        } else {
            // Если не было найдено ни одного словаря, то выдадим соответствующую ошибку
            var errorStr = "\nОтсутствуют подходящие файлы для проверки в папке \"\\dictionaries\" Файл должен иметь формат \"имя_файла_dictionary.txt\"";
            logger.error(errorStr);
            throw new NullPointerException(errorStr);
        }
        // Напечатаем строку "____"
        PrintlnBorder();

        // Вернём список именований словарей из функции
        return dictionaryNames;
    }

    /**
     * Получает именования файлов для проверки, соответствующие введённым пользователям расширениям
     * @param extensions Расширения получаемых файлов
     * @return Список именований файлов
     */
    private static List<String> GetFileNamesToCheckByInputExtensions(List<String> extensions) {

        // Зададим папку, откуда будем считывать файлы для проверки
        File folder = new File(".\\filesToCheck");

        // Считаем все файлы папки
        var files = folder.listFiles();

        // Если папки нет, то выдадим соответствующую ошибку
        if (files == null) {
            var errorStr = "\nОтсутствует папка \"..\\filesToCheck\"";
            logger.error(errorStr);
            throw new NullPointerException(errorStr);
        }
        // Если в папке нет файлов, то выдадим соответствующую ошибку
        if (files.length == 0) {
            var errorStr = "\nОтсутствуют файлы в папке \"..\\filesToCheck\"";
            logger.error(errorStr);
            throw new IllegalArgumentException(errorStr);
        }

        // Зададим лист, в который будем записывать названия файлов для проверки
        List<String> fileNamesToCheck = new ArrayList<>();
        // Зададим цикл по файлам в папке
        for (File file : files) {
            // Получим имя файла
            var fileName = file.getName();
            // Если имя файла для проверки оканчивается на одно из введённых пользователем расширений, то добавим имя файла в лист
            if (extensions.stream().anyMatch(fileName::endsWith))
                fileNamesToCheck.add(fileName);
        }

        // Напечатаем строку "____"
        PrintlnBorder();

        // Если были найдены файлы для проверки, выведем сообщение в консоль
        if (!fileNamesToCheck.isEmpty()) {
            System.out.println("Найдены файлы для проверки:");
            for (String fileName : fileNamesToCheck) {
                System.out.println(fileName);
            }
        } else {
            // Если не было найдено ни одного файла для проверки, то выдадим соответствующую ошибку
            var errorStr = GetWordsInQuotes("\nОтсутствуют файлы с расширениями ",
                    extensions, " в папке \"..\\filesToCheck\"");
            logger.error(errorStr);
            throw new NullPointerException(errorStr);
        }
        // Напечатаем строку "____"
        PrintlnBorder();

        // Вернём именования файлов для проверки из функции
        return fileNamesToCheck;
    }

    /**
     * Получает разрешения файлов для проверки, которые вводит юзер
     * @return Список разрешений файлов для проверки, которые вводит юзер
     */
    private static List<String> GetExtensionsFromConsole() {

        // Создадим лист, куда будем записывать введённые пользователем расширения
        List<String> availableExtensions = GetAvailableExtensions();

        // Создадим сканер для ввода в консоль
        Scanner scanner = new Scanner(System.in);

        // Напечатаем текст "Допустимые расширения файлов:" + доступные расширения в кавычках
        PrintlnInQuotes("Допустимые расширения файлов: ", availableExtensions);
        System.out.print("Введите расширения файлов (разделяйте пробелом если нужно несколько): ");

        // Считаем введённый пользователем текст
        String extensionsInput = scanner.nextLine();

        // Создадим лист, куда будем записывать введённые пользователем расширения (строку разделим по пробелам)
        List<String> extensions = new ArrayList<>(List.of(extensionsInput.split(" ")));

        // Создадим лист, куда будем записывать недопустимые расширения, которые ввёл юзер
        var illegalExtensions = new ArrayList<String>();
        // Запустим цикл по расширениям, введённым пользователем
        for (String fileName : extensions) {
            // Если введённое расширение нет в списке допустимых расширений, то запишем его в список недопустимых расширений
            if (!availableExtensions.contains(fileName))
                illegalExtensions.add(fileName);
        }

        // Если в списке недопустимых расширений есть хоть один элемент, то выдадим соответствующую ошибку
        if (!illegalExtensions.isEmpty()) {
            var errorStr = GetWordsInQuotes("\nНедопустимые расширения: ", illegalExtensions);
            logger.error(errorStr);
            throw new IllegalArgumentException(errorStr);
        }

        // Напечатаем текст "Выбранные расширения файлов:" + выбранные расширения в кавычках
        PrintlnInQuotes("Выбранные расширения файлов: ", extensions);

        // Закроем сканер, освободим ресурсы
        scanner.close();

        // Вернём список расширений, введённых пользователем в консоль
        return extensions;
    }

    /**
     * Печатает стартовый текст startWord + текст из слов массива words в кавычках
     * @param startWord Стартовый текст
     * @param words Слова, которые требуется обернуть в ковычки
     */
    private static void PrintlnInQuotes(String startWord, List<String> words) {

        var str = GetWordsInQuotes(startWord, words);
        // Вернём строку из функции
        System.out.println(str);
    }

    /**
     * Возвращающает стартовый текст startWord + текст из слов массива words в кавычках + текст конца строк
     * @param startWord Стартовый текст
     * @param words Слова, которые требуется обернуть в ковычки
     * @param endWord Конечный текст
     * @return Строка, состоящая из стартового текст + слов из массива в ковычках + конечный текст
     */
    private static String GetWordsInQuotes(String startWord, List<String> words, String endWord) {

        var str = GetWordsInQuotes(startWord, words);
        // Вернём строку из функции
        return str + endWord;
    }

    /**
     * Возвращает стартовый текст startWord + текст из слов массива words в кавычках
     * @param startWord Стартовый текст
     * @param words Слова, которые требуется обернуть в ковычки
     * @return Строка, состоящая из стартового текст + слов из массива в ковычках
     */
    private static String GetWordsInQuotes(String startWord, List<String> words) {

        StringBuilder result = new StringBuilder(startWord);
        var wordsLength = words.size();
        for (int i = 0; i < wordsLength; i++) {
            result.append("\"").append(words.get(i)).append("\"");
            if (i < wordsLength - 1) {
                result.append(", ");
            }
        }

        // Вернём строку из функции
        return result.toString();
    }

    /**
     * Возвращает доступные расширения файлов
     * @return Список доступных расширений файлов
     */
    private static List<String> GetAvailableExtensions() {
        return new ArrayList<>() {{
            add(".txt");
            // в процессе добавления .doc, .docx
        }};
    }

    /**
     * Функция для печати линии в консоли
     */
    private static void PrintlnBorder() {
        System.out.println("____________________________________________________________________________________");
    }
}