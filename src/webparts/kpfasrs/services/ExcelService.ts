import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import * as XLSX from 'xlsx';

// Интерфейс для результата проверки файла
export interface IFileCheckResult {
  success: boolean;
  message: string;
  filePath?: string;
  rowFound?: boolean;
  rowNumber?: number;
}

// Интерфейс для данных из списка ExportToSRS
export interface IExportToSRSItem {
  Id: number;
  Title: boolean;
  StaffMemberId: number;
  Date1: string;
  Date2: string;
  ManagerId: number;
  StaffGroupId: number;
  Condition: number;
  GroupMemberId: number;
  PathForSRSFile: string;
}

// Интерфейс для данных из списка StaffRecords
export interface IStaffRecordsItem {
  Id: number;
  Title: string;
  Date: string;
  StaffMemberId: number;
  ManagerId: number;
  StaffGroupId: number;
  StaffMember?: {
    Id: number;
    Title: string;
  };
  Checked: number;
  ExportResult: number;
  ShiftDate1: string;
  ShiftDate2: string;
  TimeForLunch: number;
  Contract: number;
  TypeOfLeaveId: number;
  TypeOfLeave?: {
    Id: number;
    Title: string;
  };
  LeaveTime: number;
  LeaveNote: string;
  LunchNote: string;
  TotalHoursNote: string;
  ReliefHours: number;
}

// URL вашего сайта SharePoint с файлами Excel
const kpfaExcelUrl: string = "https://kpfaie.sharepoint.com/sites/StaffRecordSheets";

export class ExcelService {
  private context: WebPartContext;
  
  constructor(context: WebPartContext) {
    this.context = context;
  }

  // Функция для проверки существования файла Excel и поиска строки по дате
  public async checkExcelFile(
    filePath: string, 
    selectedItem?: IExportToSRSItem, 
    groupDate?: Date
  ): Promise<IFileCheckResult> {
    try {
      // Проверяем, указан ли путь к файлу
      if (!filePath || filePath.trim() === '') {
        return {
          success: false,
          message: 'Путь к файлу не указан в выбранной записи. Проверьте поле PathForSRSFile.'
        };
      }

      // Удаляем начальный слеш из filePath, если он есть
      const cleanPath = filePath.charAt(0) === '/' ? filePath.substring(1) : filePath;

      // Формируем полный путь к файлу с правильным именем библиотеки
      const fullPath = `${kpfaExcelUrl}/Shared Documents/${cleanPath}`;
      const serverRelativePath = `/sites/StaffRecordSheets/Shared Documents/${cleanPath}`;
      console.log(`Проверка существования файла: ${fullPath}`);

      // Определяем строку для поиска на основе даты
      let dateForSearch = "";
      let dateSource = "неизвестно";
      
      if (groupDate) {
        // Используем дату группы если она передана (приоритетнее)
        try {
          dateForSearch = this.convertDateFormatSRS(groupDate);
          dateSource = "из группы";
          console.log(`Дата для поиска строки в Excel (из группы): ${dateForSearch}`);
        } catch (e) {
          console.error('Ошибка преобразования даты группы:', e);
        }
      } else if (selectedItem && selectedItem.Date1) {
        // Используем Date1 из выбранной записи как запасной вариант
        try {
          const searchDate = new Date(selectedItem.Date1);
          dateForSearch = this.convertDateFormatSRS(searchDate);
          dateSource = "из записи ExportToSRS";
          console.log(`Дата для поиска строки в Excel (из записи): ${dateForSearch}`);
        } catch (e) {
          console.error('Ошибка преобразования даты записи:', e);
        }
      }

      if (!dateForSearch) {
        return {
          success: false,
          message: 'Не удалось определить дату для поиска строки в Excel файле.'
        };
      }
      
      // API запрос для проверки существования файла
      const filePropsUrl = `${kpfaExcelUrl}/_api/web/GetFileByServerRelativeUrl('${serverRelativePath}')/Properties`;
      
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        filePropsUrl,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        if (response.status === 404) {
          // Файл не найден
          return {
            success: false,
            message: `Файл не найден: ${fullPath}\nПроверьте путь и убедитесь, что файл существует.`
          };
        } else {
          // Другая ошибка
          const errorText = await response.text();
          console.error('Error checking file:', response.status, errorText);
          return {
            success: false,
            message: `Ошибка при проверке файла: ${response.status} ${response.statusText}`
          };
        }
      }
      
      // Файл найден, теперь загружаем содержимое файла как бинарные данные
      const fileContentUrl = `${kpfaExcelUrl}/_api/web/GetFileByServerRelativeUrl('${serverRelativePath}')/$value`;
      
      const fileContentResponse: SPHttpClientResponse = await this.context.spHttpClient.get(
        fileContentUrl,
        SPHttpClient.configurations.v1
      );
      
      if (!fileContentResponse.ok) {
        return {
          success: false,
          message: `Ошибка при загрузке содержимого файла: ${fileContentResponse.status} ${fileContentResponse.statusText}`
        };
      }
      
      // Получаем содержимое файла как массив байтов
      const fileArrayBuffer = await fileContentResponse.arrayBuffer();
      
      // Теперь мы можем использовать библиотеку SheetJS для работы с Excel файлом
      try {        
        // Загружаем книгу Excel
        const workbook = XLSX.read(new Uint8Array(fileArrayBuffer), {type: 'array'});
        
        // Выводим информацию о листах в книге для отладки
        console.log('Листы в книге:', workbook.SheetNames);
        
        // Поиск листа по имени "2.Employee Data Entry"
        const targetSheetPattern = "2.Employee Data Entry";
        let targetSheetName: string | null = null;
        let findMethod = ""; // Переменная для хранения метода, которым был найден лист
        
        // Алгоритм поиска листа:
        // 1. Ищем точное совпадение
        // 2. Ищем лист, начинающийся с "2.Employee"
        // 3. Если не нашли, берем второй лист
        
        // 1. Сначала пытаемся найти точное совпадение
        for (let i = 0; i < workbook.SheetNames.length; i++) {
          const sheetName = workbook.SheetNames[i];
          if (sheetName === targetSheetPattern) {
            targetSheetName = sheetName;
            findMethod = "точное совпадение";
            break;
          }
        }
        
        // 2. Если точное совпадение не найдено, ищем по началу строки
        if (!targetSheetName) {
          for (let i = 0; i < workbook.SheetNames.length; i++) {
            const sheetName = workbook.SheetNames[i];
            if (sheetName.indexOf("2.Employee") === 0) {
              targetSheetName = sheetName;
              findMethod = "частичное совпадение (по началу имени)";
              break;
            }
          }
        }
        
        // 3. Если совпадение по имени не найдено, берем второй лист (индекс 1)
        if (!targetSheetName && workbook.SheetNames.length > 1) {
          targetSheetName = workbook.SheetNames[1];
          findMethod = "использован второй лист";
          console.log(`Лист "${targetSheetPattern}" не найден. Используем второй лист: "${targetSheetName}"`);
        }
        
        // Если нет подходящего листа, возвращаем ошибку
        if (!targetSheetName) {
          return {
            success: true,
            message: `Файл найден, но подходящий лист не найден. Доступные листы: ${workbook.SheetNames.join(", ")}. Поиск строки: "${dateForSearch}".`,
            filePath: fullPath,
            rowFound: false
          };
        }
        
        console.log(`Используем лист: "${targetSheetName}" (метод поиска: ${findMethod})`);
        
        const worksheet = workbook.Sheets[targetSheetName];
        
        // Преобразуем лист в массив объектов
        const jsonData = XLSX.utils.sheet_to_json(worksheet, {header: 1, raw: false, defval: ""});
        
        console.log(`Поиск строки с текстом "${dateForSearch}" в колонке A`);
        
        // Ищем строку, где в колонке A находится искомая дата
        let rowFound = false;
        let rowNumber = -1;
        
        for (let i = 0; i < jsonData.length; i++) {
          const row = jsonData[i] as any[];
          
          // Проверяем первую ячейку (колонка A)
          if (row && row.length > 0 && typeof row[0] === 'string') {
            const cellValue = row[0].trim();
            
            console.log(`Проверка строки ${i + 1}, значение: "${cellValue}"`);
            
            if (cellValue === dateForSearch) {
              rowFound = true;
              rowNumber = i + 1; // +1 т.к. нумерация строк в Excel начинается с 1
              console.log(`Строка найдена! Номер строки: ${rowNumber}`);
              break;
            }
          }
        }
        
        if (rowFound) {
          return {
            success: true,
            message: `1. Файл успешно найден по пути: ${fullPath}\n\n2. Поиск строки с датой "${dateForSearch}" (${dateSource}) выполнен успешно в листе "${targetSheetName}" (метод поиска листа: ${findMethod}).\n\n3. Строка найдена в позиции ${rowNumber}.`,
            filePath: fullPath,
            rowFound: true,
            rowNumber: rowNumber
          };
        } else {
          return {
            success: true,
            message: `1. Файл успешно найден по пути: ${fullPath}\n\n2. Поиск строки с датой "${dateForSearch}" (${dateSource}) выполнен в листе "${targetSheetName}" (метод поиска листа: ${findMethod}), но строка не найдена.\n\n3. Проверьте формат даты и содержимое файла Excel.`,
            filePath: fullPath,
            rowFound: false
          };
        }
        
      } catch (error) {
        console.error('Error processing Excel file:', error);
        return {
          success: true,
          message: `Файл успешно найден, но произошла ошибка при анализе содержимого Excel: ${error.message}. Поиск строки: "${dateForSearch}".`,
          filePath: fullPath,
          rowFound: false
        };
      }
      
    } catch (error) {
      console.error('Error in checkExcelFile:', error);
      return {
        success: false,
        message: `Ошибка при проверке файла: ${error.message}`
      };
    }
  }

  // Функция для преобразования даты в формат "1st of Jan" и т.д.
  public convertDateFormatSRS(inputDate: Date): string {
    try {
      // Получаем число месяца (1-31)
      const day = inputDate.getDate();
      
      // Определяем суффикс (st, nd, rd, th)
      let suffix = "th";
      if (day % 10 === 1 && day % 100 !== 11) {
        suffix = "st";
      } else if (day % 10 === 2 && day % 100 !== 12) {
        suffix = "nd";
      } else if (day % 10 === 3 && day % 100 !== 13) {
        suffix = "rd";
      }
      
      // Получаем месяц (0-11) и преобразуем в сокращение
      const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
      const month = monthNames[inputDate.getMonth()];
      
      // Формируем итоговую строку
      return `${day}${suffix} of ${month}`;
    } catch (error) {
      console.error('Error converting date:', error);
      return "Invalid Date";
    }
  }

  // Функция для проверки, является ли смена "пустой" (время 00:00)
  public isEmptyShift(record: IStaffRecordsItem): boolean {
    const shiftDate1 = record.ShiftDate1 || '';
    const shiftDate2 = record.ShiftDate2 || '';
    
    // Используем indexOf вместо endsWith для лучшей совместимости
    const isShiftDate1Empty = shiftDate1.indexOf('T00:00:00Z') === shiftDate1.length - 10 || 
                            shiftDate1.indexOf('T00:00:00') === shiftDate1.length - 9;
    
    const isShiftDate2Empty = shiftDate2.indexOf('T00:00:00Z') === shiftDate2.length - 10 || 
                            shiftDate2.indexOf('T00:00:00') === shiftDate2.length - 9;
    
    return isShiftDate1Empty && isShiftDate2Empty;
  }
}