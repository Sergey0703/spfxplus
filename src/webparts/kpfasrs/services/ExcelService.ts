import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import * as ExcelJS from 'exceljs';

// Интерфейс для результата проверки файла
export interface IFileCheckResult {
  success: boolean;
  message: string;
  filePath?: string;
  rowFound?: boolean;
  rowNumber?: number;
  cellUpdated?: boolean;
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
      const fileContentResponse: SPHttpClientResponse = await this.context.spHttpClient.get(
        `${kpfaExcelUrl}/_api/web/GetFileByServerRelativeUrl('${serverRelativePath}')/$value`,
        SPHttpClient.configurations.v1
      );
      
      if (!fileContentResponse.ok) {
        return {
          success: false,
          message: `Ошибка при загрузке содержимого файла: ${fileContentResponse.status} ${fileContentResponse.statusText}`
        };
      }
      
      // Получаем содержимое файла как ArrayBuffer
      const fileArrayBuffer = await fileContentResponse.arrayBuffer();
      
      // Теперь мы используем ExcelJS для работы с Excel файлом
      try {
        // Создаем новую книгу Excel
        const workbook = new ExcelJS.Workbook();
        
        // Загружаем содержимое файла в книгу
        await workbook.xlsx.load(fileArrayBuffer);
        
        // Выводим информацию о листах в книге для отладки
        console.log('Листы в книге:', workbook.worksheets.map(ws => ws.name));
        
        // Поиск листа по имени "2.Employee Data Entry"
        const targetSheetPattern = "2.Employee Data Entry";
        let targetWorksheet: ExcelJS.Worksheet | undefined;
        let findMethod = ""; // Переменная для хранения метода, которым был найден лист
        
        // Алгоритм поиска листа:
        // 1. Ищем точное совпадение
        // 2. Ищем лист, начинающийся с "2.Employee"
        // 3. Если не нашли, берем второй лист
        
        // 1. Сначала пытаемся найти точное совпадение
        targetWorksheet = workbook.getWorksheet(targetSheetPattern);
        if (targetWorksheet) {
          findMethod = "точное совпадение";
        }
        
        // 2. Если точное совпадение не найдено, ищем по началу строки
        if (!targetWorksheet) {
          for (const worksheet of workbook.worksheets) {
            if (worksheet.name.indexOf("2.Employee") === 0) {
              targetWorksheet = worksheet;
              findMethod = "частичное совпадение (по началу имени)";
              break;
            }
          }
        }
        
        // 3. Если совпадение по имени не найдено, берем второй лист (индекс 1)
        if (!targetWorksheet && workbook.worksheets.length > 1) {
          targetWorksheet = workbook.worksheets[1];
          findMethod = "использован второй лист";
          console.log(`Лист "${targetSheetPattern}" не найден. Используем второй лист: "${targetWorksheet.name}"`);
        }
        
        // Если нет подходящего листа, возвращаем ошибку
        if (!targetWorksheet) {
          return {
            success: true,
            message: `Файл найден, но подходящий лист не найден. Доступные листы: ${workbook.worksheets.map(ws => ws.name).join(", ")}. Поиск строки: "${dateForSearch}".`,
            filePath: fullPath,
            rowFound: false
          };
        }
        
        const targetSheetName = targetWorksheet.name;
        console.log(`Используем лист: "${targetSheetName}" (метод поиска: ${findMethod})`);
        
        // Ищем строку, где в колонке A находится искомая дата
        let rowFound = false;
        let rowNumber = -1;
        
        // Перебираем строки в листе
        targetWorksheet.eachRow({ includeEmpty: false }, (row, rowNum) => {
          // Получаем значение в первой ячейке (колонка A)
          const cellValue = row.getCell(1).text;
          
          console.log(`Проверка строки ${rowNum}, значение: "${cellValue}"`);
          
          if (cellValue === dateForSearch) {
            rowFound = true;
            rowNumber = rowNum;
            console.log(`Строка найдена! Номер строки: ${rowNumber}`);
            // Прерываем итерацию, так как строка найдена
            return false;
          }
        });
        
        if (rowFound) {
          // Если строка найдена, записываем время "20:20" в ячейку B этой строки
          try {
            // Получаем ячейку B в найденной строке
            const cell = targetWorksheet.getRow(rowNumber).getCell(2);
            
            // Устанавливаем значение ячейки
            cell.value = "20:20";
            
            // Сохраняем изменения обратно в файл
            const updatedContent = await workbook.xlsx.writeBuffer();
            
            // Отправляем обновленный файл на SharePoint
            await this.updateExcelFile(serverRelativePath, updatedContent);
            
            return {
              success: true,
              message: `1. Файл успешно найден по пути: ${fullPath}\n\n2. Поиск строки с датой "${dateForSearch}" (${dateSource}) выполнен успешно в листе "${targetSheetName}" (метод поиска листа: ${findMethod}).\n\n3. Строка найдена в позиции ${rowNumber}.\n\n4. Значение "20:20" записано в ячейку B${rowNumber} и файл успешно обновлен.`,
              filePath: fullPath,
              rowFound: true,
              rowNumber: rowNumber,
              cellUpdated: true
            };
          } catch (error) {
            console.error('Error updating cell:', error);
            return {
              success: true,
              message: `1. Файл успешно найден по пути: ${fullPath}\n\n2. Поиск строки с датой "${dateForSearch}" (${dateSource}) выполнен успешно в листе "${targetSheetName}" (метод поиска листа: ${findMethod}).\n\n3. Строка найдена в позиции ${rowNumber}.\n\n4. Не удалось записать значение в ячейку B${rowNumber}: ${error.message}`,
              filePath: fullPath,
              rowFound: true,
              rowNumber: rowNumber,
              cellUpdated: false
            };
          }
        } else {
          return {
            success: true,
            message: `1. Файл успешно найден по пути: ${fullPath}\n\n2. Поиск строки с датой "${dateForSearch}" (${dateSource}) выполнен в листе "${targetSheetName}" (метод поиска листа: ${findMethod}), но строка не найдена.\n\n3. Проверьте формат даты и содержимое файла Excel.`,
            filePath: fullPath,
            rowFound: false,
            cellUpdated: false
          };
        }
        
      } catch (error) {
        console.error('Error processing Excel file:', error);
        return {
          success: true,
          message: `Файл успешно найден, но произошла ошибка при анализе содержимого Excel: ${error.message}. Поиск строки: "${dateForSearch}".`,
          filePath: fullPath,
          rowFound: false,
          cellUpdated: false
        };
      }
      
    } catch (error) {
      console.error('Error in checkExcelFile:', error);
      return {
        success: false,
        message: `Ошибка при проверке файла: ${error.message}`,
        cellUpdated: false
      };
    }
  }

  // Функция для обновления файла Excel на SharePoint с использованием современных методов
  private async updateExcelFile(serverRelativePath: string, fileData: ArrayBuffer): Promise<void> {
    try {
      // Используем метод $value для установки содержимого файла
      const url = `${kpfaExcelUrl}/_api/web/GetFileByServerRelativeUrl('${serverRelativePath}')/$value`;
      
      // Получаем токен запроса для редактирования
      const digestResponse: SPHttpClientResponse = await this.context.spHttpClient.post(
        `${kpfaExcelUrl}/_api/contextinfo`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata'
          }
        }
      );
      
      if (!digestResponse.ok) {
        throw new Error(`Ошибка получения токена: ${digestResponse.status}`);
      }
      
      const digestValue = (await digestResponse.json()).FormDigestValue;
      
      // Обновляем файл
      const updateResponse: SPHttpClientResponse = await this.context.spHttpClient.post(
        url,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-Type': 'application/octet-stream',
            'X-HTTP-Method': 'PUT',
            'X-RequestDigest': digestValue
          },
          body: fileData
        }
      );
      
      if (!updateResponse.ok) {
        const errorText = await updateResponse.text();
        console.error('Error updating file:', updateResponse.status, errorText);
        throw new Error(`Ошибка обновления файла: ${updateResponse.status} ${updateResponse.statusText}`);
      }
      
      console.log('Файл успешно обновлен.');
    } catch (error) {
      console.error('Error in updateExcelFile:', error);
      throw error;
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