import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

// Интерфейс для результата проверки файла
export interface IFileCheckResult {
  success: boolean;
  message: string;
  filePath?: string;
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

      // API запрос для проверки существования файла
      // Используем подход с получением свойств файла вместо его содержимого
      const filePropsUrl = `${kpfaExcelUrl}/_api/web/GetFileByServerRelativeUrl('${serverRelativePath}')/Properties`;
      
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        filePropsUrl,
        SPHttpClient.configurations.v1
      );

      if (response.ok) {
        // Файл найден
        // Теперь мы должны определить, какую строку искать (на основе даты)
        let dateForSearch = "";
        
        if (groupDate) {
          // Используем дату группы если она передана (приоритетнее)
          try {
            dateForSearch = this.convertDateFormatSRS(groupDate);
            console.log(`Дата для поиска строки в Excel (из группы): ${dateForSearch}`);
          } catch (e) {
            console.error('Ошибка преобразования даты группы:', e);
          }
        } else if (selectedItem && selectedItem.Date1) {
          // Используем Date1 из выбранной записи как запасной вариант
          try {
            const searchDate = new Date(selectedItem.Date1);
            dateForSearch = this.convertDateFormatSRS(searchDate);
            console.log(`Дата для поиска строки в Excel (из записи): ${dateForSearch}`);
          } catch (e) {
            console.error('Ошибка преобразования даты записи:', e);
          }
        }
        
        // Пока что мы только проверили наличие файла, но не выполняли поиск строки
        // В реальном коде здесь будет логика поиска строки в Excel
        // Пока используем заглушку: строка "не найдена"
        const foundStringStatus = "Строка не найдена. Потребуется реализация поиска внутри Excel-файла.";
        
        return {
          success: true,
          message: `1. Файл успешно найден по пути: ${fullPath}\n\n2. Строка для поиска: "${dateForSearch}"\n\n3. Результат поиска строки: ${foundStringStatus}`,
          filePath: fullPath
        };
      } else if (response.status === 404) {
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