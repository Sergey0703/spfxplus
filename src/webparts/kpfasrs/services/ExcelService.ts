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

      // Пробуем разные варианты путей к файлу
      const possiblePaths = [
        // Путь 1: Как указано в PathForSRSFile
        cleanPath,
        // Путь 2: Обрезанный путь (без указания Shared Documents)
        cleanPath.replace('Shared Documents/', ''),
        // Путь 3: Путь непосредственно к файлу без структуры каталогов
        cleanPath.split('/').pop() || '',
        // Путь 4: Альтернативный путь в другой структуре папок
        `Kilcummin Residential Services/Lohan Lodge/SRSExport/Sinead Twomey 2024Test.xlsx`
      ];

      // Формируем полные серверные пути для проверки
      const serverPaths = possiblePaths.map(path => {
        if (path) {
          return `/sites/StaffRecordSheets/Shared Documents/${path}`;
        }
        return ''; // Пропускаем пустые пути
      }).filter(p => p !== '');

      console.log('Проверяем возможные пути к файлу:');
      serverPaths.forEach(p => console.log(` - ${p}`));

      // Проверяем каждый путь по очереди, пока не найдем существующий файл
      let fileExists = false;
      let serverRelativePath = '';
      let fullPath = '';

      for (const path of serverPaths) {
        try {
          console.log(`Проверка существования файла: ${path}`);
          const filePropsUrl = `${kpfaExcelUrl}/_api/web/GetFileByServerRelativeUrl('${path}')/Properties`;
          
          const response: SPHttpClientResponse = await this.context.spHttpClient.get(
            filePropsUrl,
            SPHttpClient.configurations.v1
          );

          if (response.ok) {
            fileExists = true;
            serverRelativePath = path;
            fullPath = `${kpfaExcelUrl}${path.replace('/sites/StaffRecordSheets', '')}`;
            console.log(`Файл найден по пути: ${fullPath}`);
            break;
          }
        } catch (error) {
          console.log(`Файл не найден по пути: ${path}`, error);
        }
      }

      if (!fileExists) {
        return {
          success: false,
          message: `Файл не найден ни по одному из проверенных путей:\n${serverPaths.join('\n')}\n\nПроверьте путь и убедитесь, что файл существует.`
        };
      }

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
        
        // Вывод первых 20 значений из колонки A для отладки
        const firstColumnValues: string[] = [];
        
        // Переменные для конкретной строки 1008
        let cellA1008Value = "";
        let cellA1008Details = "";
        let cellA1008HexDump = "";
        
        targetWorksheet.eachRow({ includeEmpty: false }, (row, rowNum) => {
          // Сохраняем первые 20 значений для отладки
          if (rowNum <= 20) {
            const cellValue = row.getCell(1).text;
            firstColumnValues.push(`Строка ${rowNum}: "${cellValue}"`);
          }
          
          // Проверяем строку 1008
          if (rowNum === 1008) {
            const cell = row.getCell(1);
            cellA1008Value = cell.text || "";
            
            // Подробная информация о ячейке
            cellA1008Details = `Тип: ${typeof cellA1008Value}, Длина: ${cellA1008Value.length}`;
            
            // Создаем hex-дамп для сравнения каждого символа
            let hexDump = "";
            for (let i = 0; i < cellA1008Value.length; i++) {
              const charCode = cellA1008Value.charCodeAt(i);
              hexDump += `${cellA1008Value[i]} (${charCode.toString(16)}), `;
            }
            cellA1008HexDump = hexDump;
            
            console.log(`Ячейка A1008 найдена! Значение: "${cellA1008Value}"`);
            console.log(`Детали ячейки A1008: ${cellA1008Details}`);
            console.log(`Hex-дамп ячейки A1008: ${cellA1008HexDump}`);
          }
        });
        console.log('Первые 20 значений в колонке A:', firstColumnValues.join('\n'));
        
        // Ищем строку, где в колонке A находится искомая дата
        let rowFound = false;
        let rowNumber = -1;
        
        // Нормализуем значение для поиска
        const normalizedSearch = dateForSearch.trim().toLowerCase();
        const shortFormatSearch = normalizedSearch.replace(" of ", " ");
        
        // Создаем hex-дамп для строки поиска для сравнения
        let searchHexDump = "";
        for (let i = 0; i < dateForSearch.length; i++) {
            const charCode = dateForSearch.charCodeAt(i);
            searchHexDump += `${dateForSearch[i]} (${charCode.toString(16)}), `;
        }
        console.log(`Строка поиска: "${dateForSearch}"`);
        console.log(`Тип: ${typeof dateForSearch}, Длина: ${dateForSearch.length}`);
        console.log(`Hex-дамп строки поиска: ${searchHexDump}`);
        
        // Перебираем строки в листе
        targetWorksheet.eachRow({ includeEmpty: false }, (row, rowNum) => {
          // Получаем значение в первой ячейке (колонка A)
          const cellValue = row.getCell(1).text;
          
          if (!cellValue) return; // Пропускаем пустые ячейки
          
          console.log(`Проверка строки ${rowNum}, значение: "${cellValue}"`);
          
          // Нормализуем значение ячейки
          const normalizedCell = cellValue.trim().toLowerCase();
          const shortFormatCell = normalizedCell.replace(" of ", " ");
          
          // Если это строка 1008, дополнительно выводим информацию для сравнения
          if (rowNum === 1008) {
            console.log(`Сравнение в строке 1008:`);
            console.log(`- Значение ячейки: "${cellValue}"`);
            console.log(`- Значение ячейки после нормализации: "${normalizedCell}"`);
            console.log(`- Искомая строка: "${dateForSearch}"`);
            console.log(`- Искомая строка после нормализации: "${normalizedSearch}"`);
            console.log(`- Равны ли значения без учета регистра: ${normalizedCell === normalizedSearch}`);
            
            // Проверяем посимвольно для выявления различий
            console.log(`- Посимвольное сравнение:`);
            const maxLength = Math.max(normalizedCell.length, normalizedSearch.length);
            for (let i = 0; i < maxLength; i++) {
              const cellChar = i < normalizedCell.length ? normalizedCell.charCodeAt(i) : -1;
              const searchChar = i < normalizedSearch.length ? normalizedSearch.charCodeAt(i) : -1;
              console.log(`  Позиция ${i}: ${normalizedCell[i] || ''} (${cellChar}) vs ${normalizedSearch[i] || ''} (${searchChar}) - ${cellChar === searchChar ? 'совпадает' : 'отличается'}`);
            }
          }
          
          // Проверяем различные варианты написания даты
          if (normalizedCell === normalizedSearch) {
            rowFound = true;
            rowNumber = rowNum;
            console.log(`Строка найдена! Номер строки: ${rowNumber}, точное совпадение`);
            return false;
          }
          
          // Проверяем сокращенный формат (без "of")
          if (shortFormatCell === shortFormatSearch || 
              shortFormatCell === normalizedSearch || 
              normalizedCell === shortFormatSearch) {
            rowFound = true;
            rowNumber = rowNum;
            console.log(`Строка найдена! Номер строки: ${rowNumber}, альтернативный формат`);
            return false;
          }
          
          // Проверяем на совпадение день и месяц (могут быть разные форматы суффиксов)
          if (rowNum === 1008) {
            // Проверка на совпадение по шаблону "1(st|nd|rd|th) of Dec" без учета суффиксов
            const regexCellMatch = normalizedCell.match(/(\d+)(?:st|nd|rd|th)?\s+(?:of\s+)?([a-z]{3})/i);
            const regexSearchMatch = normalizedSearch.match(/(\d+)(?:st|nd|rd|th)?\s+(?:of\s+)?([a-z]{3})/i);
            
            if (regexCellMatch && regexSearchMatch) {
              const cellDay = regexCellMatch[1];
              const cellMonth = regexCellMatch[2].toLowerCase();
              const searchDay = regexSearchMatch[1];
              const searchMonth = regexSearchMatch[2].toLowerCase();
              
              console.log(`- Извлечение по шаблону для строки 1008:`);
              console.log(`  Ячейка: День=${cellDay}, Месяц=${cellMonth}`);
              console.log(`  Поиск: День=${searchDay}, Месяц=${searchMonth}`);
              
              if (cellDay === searchDay && cellMonth === searchMonth) {
                console.log(`  Дни и месяцы совпадают!`);
              }
            }
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
            try {
              await this.updateExcelFile(serverRelativePath, updatedContent);
              return {
                success: true,
                message: `1. Файл успешно найден по пути: ${fullPath}\n\n2. Поиск строки с датой "${dateForSearch}" (${dateSource}) выполнен успешно в листе "${targetSheetName}" (метод поиска листа: ${findMethod}).\n\n3. Строка найдена в позиции ${rowNumber}.\n\n4. Значение "20:20" записано в ячейку B${rowNumber} и файл успешно обновлен.`,
                filePath: fullPath,
                rowFound: true,
                rowNumber: rowNumber,
                cellUpdated: true
              };
            } catch (updateError) {
              // Проверяем, содержит ли сообщение ошибки информацию о блокировке файла
              const errorMessage = updateError.message || '';
              if (errorMessage.includes('423') || errorMessage.includes('заблокирован')) {
                return {
                  success: true,
                  message: `1. Файл успешно найден по пути: ${fullPath}\n\n2. Поиск строки с датой "${dateForSearch}" (${dateSource}) выполнен успешно в листе "${targetSheetName}" (метод поиска листа: ${findMethod}).\n\n3. Строка найдена в позиции ${rowNumber}.\n\n4. Не удалось записать значение в ячейку B${rowNumber}: файл заблокирован. Возможно, файл Excel открыт для редактирования другим пользователем. Пожалуйста, убедитесь, что файл закрыт всеми пользователями и повторите попытку.`,
                  filePath: fullPath,
                  rowFound: true,
                  rowNumber: rowNumber,
                  cellUpdated: false
                };
              } else {
                return {
                  success: true,
                  message: `1. Файл успешно найден по пути: ${fullPath}\n\n2. Поиск строки с датой "${dateForSearch}" (${dateSource}) выполнен успешно в листе "${targetSheetName}" (метод поиска листа: ${findMethod}).\n\n3. Строка найдена в позиции ${rowNumber}.\n\n4. Не удалось записать значение в ячейку B${rowNumber}: ${updateError.message}`,
                  filePath: fullPath,
                  rowFound: true,
                  rowNumber: rowNumber,
                  cellUpdated: false
                };
              }
            }
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
          // Проверяем, есть ли ячейка A1008 и сравниваем ее со строкой поиска
          let comparisonInfo = "";
          if (cellA1008Value) {
            comparisonInfo = `\n\n5. Сравнение ячейки A1008 с искомой строкой:\n` +
              `   - Значение в A1008: "${cellA1008Value}"\n` +
              `   - Искомая строка: "${dateForSearch}"\n` +
              `   - Значение ячейки в нижнем регистре: "${cellA1008Value.toLowerCase()}"\n` +
              `   - Искомая строка в нижнем регистре: "${dateForSearch.toLowerCase()}"\n` +
              `   - Значение ячейки без пробелов: "${cellA1008Value.trim()}"\n` +
              `   - Искомая строка без пробелов: "${dateForSearch.trim()}"\n`;
          }
          
          return {
            success: true,
            message: `1. Файл успешно найден по пути: ${fullPath}\n\n2. Поиск строки с датой "${dateForSearch}" (${dateSource}) выполнен в листе "${targetSheetName}" (метод поиска листа: ${findMethod}), но строка не найдена.\n\n3. Проверьте формат даты и содержимое файла Excel.\n\n4. Значения первых ячеек: ${firstColumnValues.slice(0, 5).join(', ')}...${comparisonInfo}`,
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
      console.log(`Обновление файла по пути: ${serverRelativePath}`);
      
      // Используем метод $value для установки содержимого файла
      const url = `${kpfaExcelUrl}/_api/web/GetFileByServerRelativeUrl('${serverRelativePath}')/$value`;
      
      // Метод PUT без запроса токена
      const updateResponse: SPHttpClientResponse = await this.context.spHttpClient.fetch(
        url,
        SPHttpClient.configurations.v1,
        {
          method: 'PUT',
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-Type': 'application/octet-stream'
          },
          body: fileData
        }
      );
      
      if (!updateResponse.ok) {
        const errorText = await updateResponse.text();
        console.error('Error updating file:', updateResponse.status, errorText);
        
        // Специальная обработка ошибки 423 (Locked)
        if (updateResponse.status === 423) {
          throw new Error(`Файл заблокирован (код 423). Возможно, файл открыт для редактирования другим пользователем. Пожалуйста, убедитесь, что файл Excel закрыт всеми пользователями и повторите попытку.`);
        } else {
          throw new Error(`Ошибка обновления файла: ${updateResponse.status} ${updateResponse.statusText}`);
        }
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
      
      // Формируем итоговую строку и удаляем лишние пробелы
      const result = `${day}${suffix} of ${month}`;
      return result.trim();
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