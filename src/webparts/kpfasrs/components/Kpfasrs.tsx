import * as React from 'react';
import { useState, useEffect } from 'react';
import styles from './Kpfasrs.module.scss';
import { IKpfasrsProps } from './IKpfasrsProps';
import { 
  DetailsList, 
  SelectionMode, 
  IColumn,
  MessageBar,
  MessageBarType,
  Selection,
  Spinner,
  SpinnerSize,
  IGroup,
  GroupHeader,
  IGroupHeaderProps,
  PrimaryButton,
  DefaultButton,
  Dialog,
  DialogType,
  DialogFooter,
  IObjectWithKey
} from '@fluentui/react';
import { SharePointService } from '../services/SharePointService';
import { ExcelService, IExportToSRSItem, IStaffRecordsItem, IFileCheckResult } from '../services/ExcelService';
import { DataUtils } from '../services/DataUtils';

// Расширяем интерфейс IExportToSRSItem, чтобы он соответствовал IObjectWithKey
interface IExportToSRSItemWithKey extends IExportToSRSItem, IObjectWithKey {
  key: string; // Добавляем свойство key для соответствия IObjectWithKey
}

const Kpfasrs: React.FC<IKpfasrsProps> = (props) => {
  // Состояния для ExportToSRS
  const [items, setItems] = useState<IExportToSRSItemWithKey[]>([]);
  const [isLoading, setIsLoading] = useState<boolean>(false);
  
  // Состояния для StaffRecords
  const [staffRecords, setStaffRecords] = useState<IStaffRecordsItem[]>([]);
  const [filteredStaffRecords, setFilteredStaffRecords] = useState<IStaffRecordsItem[]>([]);
  const [staffRecordsGroups, setStaffRecordsGroups] = useState<IGroup[]>([]);
  const [isLoadingStaffRecords, setIsLoadingStaffRecords] = useState<boolean>(false);
  const [selectedItem, setSelectedItem] = useState<IExportToSRSItemWithKey | null>(null);
  
  // Состояния для экспорта и диалогов
  const [isExporting, setIsExporting] = useState<boolean>(false);
  const [showDialog, setShowDialog] = useState<boolean>(false);
  const [dialogMessage, setDialogMessage] = useState<string>('');
  const [dialogTitle, setDialogTitle] = useState<string>('');
  const [isErrorDialog, setIsErrorDialog] = useState<boolean>(false);
  const [lastSearchResult, setLastSearchResult] = useState<IFileCheckResult | null>(null);
  
  // Состояния для автообработки
  const [isAutoProcessing, setIsAutoProcessing] = useState<boolean>(false);
  const [currentAutoIndex, setCurrentAutoIndex] = useState<number>(-1);
  const [processingResults, setProcessingResults] = useState<string[]>([]);
  
  // Общие состояния
  const [error, setError] = useState<string | null>(null);
  const [debugInfo, setDebugInfo] = useState<string | null>(null);

  // Инициализация сервисов
  const sharePointService = React.useMemo(() => new SharePointService(props.context), [props.context]);
  const excelService = React.useMemo(() => new ExcelService(props.context), [props.context]);
  const dataUtils = React.useMemo(() => new DataUtils(excelService), [excelService]);

  // Отслеживание изменения выбранного элемента
  useEffect(() => {
    if (isAutoProcessing && selectedItem) {
      console.log(`Начинаем обработку элемента ID: ${selectedItem.Id}`);
      // Обрабатываем выбранный элемент - загружаем и фильтруем записи
      handleFilterStaffRecords(selectedItem);
    }
  }, [selectedItem, isAutoProcessing]); // Зависимость только от selectedItem и isAutoProcessing

  // Отдельный useEffect для обработки отфильтрованных записей
  useEffect(() => {
    if (isAutoProcessing && selectedItem && filteredStaffRecords.length > 0 && staffRecordsGroups.length > 0) {
      console.log(`Запускаем автоэкспорт для элемента ID: ${selectedItem.Id}, найдено групп: ${staffRecordsGroups.length}`);
      // Запускаем экспорт только если есть отфильтрованные записи и группы
      handleAutoExport();
    }
  }, [filteredStaffRecords, staffRecordsGroups, isAutoProcessing, selectedItem]); // Зависит от результатов фильтрации

  // Определение колонок для ExportToSRS
  const columns: IColumn[] = [
    { key: 'id', name: 'ID', fieldName: 'Id', minWidth: 50, maxWidth: 50 },
    { key: 'title', name: 'Title', fieldName: 'Title', minWidth: 100, maxWidth: 150, onRender: (item) => (item.Title ? 'Yes' : 'No') },
    { key: 'staffMember', name: 'Staff Member', fieldName: 'StaffMemberId', minWidth: 100, maxWidth: 150 },
    { key: 'date1', name: 'Date 1', fieldName: 'Date1', minWidth: 100, maxWidth: 150 },
    { key: 'date2', name: 'Date 2', fieldName: 'Date2', minWidth: 100, maxWidth: 150 },
    { key: 'manager', name: 'Manager', fieldName: 'ManagerId', minWidth: 100, maxWidth: 150 },
    { key: 'staffGroup', name: 'Staff Group', fieldName: 'StaffGroupId', minWidth: 100, maxWidth: 150 },
    { key: 'condition', name: 'Condition', fieldName: 'Condition', minWidth: 100, maxWidth: 150 },
    { key: 'groupMember', name: 'Group Member', fieldName: 'GroupMemberId', minWidth: 100, maxWidth: 150 },
    { key: 'pathForSRSFile', name: 'Path For SRS File', fieldName: 'PathForSRSFile', minWidth: 200, maxWidth: 300 },
    { key: 'email', name: 'Email', fieldName: 'email', minWidth: 200, maxWidth: 300 }
  ];
  
  // Определение колонок для StaffRecords с добавленными полями
  const staffRecordsColumns: IColumn[] = [
    { key: 'id', name: 'ID', fieldName: 'Id', minWidth: 50, maxWidth: 50 },
    { key: 'title', name: 'Title', fieldName: 'Title', minWidth: 100, maxWidth: 150 },
    { key: 'date', name: 'Date', fieldName: 'Date', minWidth: 100, maxWidth: 150 },
    { key: 'shiftDate1', name: 'Shift Start', fieldName: 'ShiftDate1', minWidth: 120, maxWidth: 150 },
    { key: 'shiftDate2', name: 'Shift End', fieldName: 'ShiftDate2', minWidth: 120, maxWidth: 150 },
    { key: 'staffMember', name: 'Staff Member', fieldName: 'StaffMemberId', minWidth: 100, maxWidth: 150 },
    { key: 'manager', name: 'Manager', fieldName: 'ManagerId', minWidth: 100, maxWidth: 150 },
    { key: 'staffGroup', name: 'Staff Group', fieldName: 'StaffGroupId', minWidth: 100, maxWidth: 150 },
    { key: 'typeOfLeave', name: 'Type Of Leave', fieldName: 'TypeOfLeaveId', minWidth: 100, maxWidth: 150 },
    { key: 'contract', name: 'Contract', fieldName: 'Contract', minWidth: 80, maxWidth: 100 },
    { key: 'timeForLunch', name: 'Time For Lunch', fieldName: 'TimeForLunch', minWidth: 100, maxWidth: 120 },
    { key: 'leaveTime', name: 'Leave Time', fieldName: 'LeaveTime', minWidth: 80, maxWidth: 100 },
    { key: 'reliefHours', name: 'Relief Hours', fieldName: 'ReliefHours', minWidth: 80, maxWidth: 100 },
    { key: 'leaveNote', name: 'Leave Note', fieldName: 'LeaveNote', minWidth: 120, maxWidth: 200 },
    { key: 'lunchNote', name: 'Lunch Note', fieldName: 'LunchNote', minWidth: 120, maxWidth: 200 },
    { key: 'totalHoursNote', name: 'Total Hours Note', fieldName: 'TotalHoursNote', minWidth: 120, maxWidth: 200 },
    { key: 'checked', name: 'Checked', fieldName: 'Checked', minWidth: 80, maxWidth: 100 },
    { key: 'exportResult', name: 'Export Result', fieldName: 'ExportResult', minWidth: 100, maxWidth: 120 }
  ];

  // Создаем объект Selection для DetailsList
  const selection = new Selection({
    onSelectionChanged: () => {
      // Срабатывает только при ручном выборе в интерфейсе
      if (isAutoProcessing) return; // Пропускаем, если идет автообработка

      const selectedItems = selection.getSelection() as IExportToSRSItemWithKey[];
      if (selectedItems.length > 0) {
        const selectedExportItem = selectedItems[0];
        setSelectedItem(selectedExportItem);
        handleFilterStaffRecords(selectedExportItem);
      } else {
        setSelectedItem(null);
        setFilteredStaffRecords([]);
        setStaffRecordsGroups([]);
      }
    }
  });
  
  // Функция для форматирования HTML-содержимого письма
  const formatEmailBody = (title: string, message: string): string => {
    return `
      <html>
        <head>
          <style>
            body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 0; padding: 20px; color: #333; }
            .container { max-width: 600px; margin: 0 auto; border: 1px solid #ddd; border-radius: 4px; padding: 20px; }
            h2 { color: #0078d4; border-bottom: 1px solid #eee; padding-bottom: 10px; }
            .results { white-space: pre-line; background-color: #f9f9f9; padding: 15px; border-radius: 4px; font-family: monospace; }
            .footer { margin-top: 20px; font-size: 12px; color: #666; border-top: 1px solid #eee; padding-top: 10px; }
          </style>
        </head>
        <body>
          <div class="container">
            <h2>${title}</h2>
            <div class="results">${message.replace(/\n/g, '<br>')}</div>
            <div class="footer">
              Это автоматическое уведомление. Пожалуйста, не отвечайте на это письмо.
            </div>
          </div>
        </body>
      </html>
    `;
  };

  // Новая функция для загрузки данных из ExportToSRS при нажатии на кнопку
  const handleLoadExportToSRSItems = async (): Promise<void> => {
    try {
      setIsLoading(true);
      setError(null);
      
      const loadedItems = await sharePointService.getExportToSRSItems();
      
      // Преобразуем IExportToSRSItem в IExportToSRSItemWithKey
      const itemsWithKey: IExportToSRSItemWithKey[] = loadedItems.map(item => ({
        ...item,
        key: `item-${item.Id}` // Добавляем уникальный ключ
      }));
      
      setItems(itemsWithKey);
      
      // Автоматическое начало обработки после загрузки данных
      if (itemsWithKey && itemsWithKey.length > 0) {
        // Запускаем автоматическую обработку
        startAutoProcessing(itemsWithKey);
      }
    } catch (error) {
      console.error('Error in loadExportToSRSItems:', error);
      setError(`Ошибка: ${error.message}`);
    } finally {
      setIsLoading(false);
    }
  };

  // Функция для запуска автоматической обработки всех элементов
  const startAutoProcessing = (loadedItems: IExportToSRSItemWithKey[]): void => {
    // Фильтруем элементы с непустым PathForSRSFile
    const validItems = loadedItems.filter(item => item.PathForSRSFile && item.PathForSRSFile.trim() !== '');
    
    if (validItems.length === 0) {
      setError('Нет элементов с указанным путем к файлу (PathForSRSFile)');
      return;
    }
    
    // Начинаем автоматическую обработку
    setIsAutoProcessing(true);
    setProcessingResults([]);
    setCurrentAutoIndex(0);
    
    // Сбрасываем предыдущие результаты и состояния
    setFilteredStaffRecords([]);
    setStaffRecordsGroups([]);
    setLastSearchResult(null);
    
    // Выбираем первый элемент для обработки
    selection.setItems(loadedItems);
    selection.setKeySelected(validItems[0].key, true, false);
    
    // Небольшая пауза перед установкой selectedItem
    setTimeout(() => {
      // Устанавливаем selectedItem напрямую, чтобы инициировать обработку
      setSelectedItem(validItems[0]);
      console.log(`Начинаем автоматическую обработку ${validItems.length} элементов, первый ID: ${validItems[0].Id}`);
    }, 500); // 500 мс пауза
  };

  // Функция для автоматического экспорта данных с повторными попытками
  const handleAutoExport = async (): Promise<void> => {
    if (!selectedItem) {
      finishCurrentItemProcessing('Не выбрана запись для экспорта');
      moveToNextItem();
      return;
    }
    
    // Проверяем, есть ли группы
    if (staffRecordsGroups.length === 0) {
      finishCurrentItemProcessing(`ID ${selectedItem.Id} - Нет доступных дней для экспорта`);
      moveToNextItem();
      return;
    }
    
    try {
      setIsExporting(true);
      
      // Получаем ключ (дату) первой группы
      const firstGroupKey = staffRecordsGroups[0].key as string;
      console.log(`Автоматический экспорт для первой группы с ключом: ${firstGroupKey}`);
      
      // Получаем путь к файлу из выбранной записи
      const filePath = selectedItem.PathForSRSFile;
      
      // Используем дату из первой группы
      let groupDate: Date | undefined;
      try {
        groupDate = new Date(firstGroupKey);
      } catch (e) {
        console.error('Не удалось преобразовать ключ группы в дату:', firstGroupKey, e);
      }
      
      // Максимальное количество попыток
      const maxAttempts = 3;
      let currentAttempt = 1;
      let lastError = null;
      let checkResult = null;
      
      // Делаем несколько попыток в случае ошибки блокировки файла
      while (currentAttempt <= maxAttempts) {
        try {
          console.log(`Попытка ${currentAttempt} из ${maxAttempts} экспорта для ID ${selectedItem.Id}, дата ${firstGroupKey}`);
          
          // Проверяем существование файла и выполняем экспорт
          checkResult = await excelService.checkExcelFile(filePath, selectedItem, groupDate);
          
          // Если успешно записали значение или получили ошибку, которая не связана с блокировкой,
          // прерываем цикл
          if (checkResult.success && (checkResult.cellUpdated || !checkResult.rowFound)) {
            break;
          }
          
          // Если файл найден, строка найдена, но значение не удалось записать из-за блокировки
          if (checkResult.success && checkResult.rowFound && !checkResult.cellUpdated) {
            const errorMsg = checkResult.message || "";
            // Проверяем, есть ли признаки блокировки файла
            if (errorMsg.includes("заблокирован") || errorMsg.includes("423") || errorMsg.includes("locked")) {
              console.log(`Файл заблокирован. Ожидаем и повторяем попытку ${currentAttempt}`);
              // Ждем некоторое время перед следующей попыткой
              await new Promise(resolve => setTimeout(resolve, 3000)); // 3 секунды
              currentAttempt++;
              lastError = new Error("Файл заблокирован");
            } else {
              // Другая ошибка, прерываем цикл
              break;
            }
          } else {
            // Успешное выполнение или ошибка, не связанная с блокировкой
            break;
          }
        } catch (error) {
          // Проверяем, связана ли ошибка с блокировкой файла
          if (error.message && (
            error.message.includes("заблокирован") || 
            error.message.includes("423") || 
            error.message.includes("locked")
          )) {
            console.log(`Ошибка блокировки. Ожидаем и повторяем попытку ${currentAttempt}`);
            // Ждем некоторое время перед следующей попыткой
            await new Promise(resolve => setTimeout(resolve, 3000)); // 3 секунды
            currentAttempt++;
            lastError = error;
          } else {
            // Другая ошибка, прерываем цикл
            throw error;
          }
        }
      }
      
      // Если после всех попыток мы все еще получаем ошибку блокировки
      if (currentAttempt > maxAttempts && lastError) {
        throw lastError;
      }
      
      // Сохраняем результат поиска
      if (checkResult) {
        setLastSearchResult(checkResult);
      }
      
      // Формируем сообщение о результате
      let resultMessage = `ID ${selectedItem.Id}, Дата: ${firstGroupKey}`;
      
      if (checkResult && checkResult.success) {
        if (checkResult.rowFound) {
          resultMessage += ` - Успешно: найдена строка ${checkResult.rowNumber}`;
          if (checkResult.cellUpdated) {
            resultMessage += ', значение обновлено';
          } else {
            resultMessage += ', но не удалось обновить значение';
          }
        } else {
          resultMessage += ` - Строка не найдена`;
        }
      } else if (checkResult) {
        resultMessage += ` - Ошибка: ${checkResult.message.substring(0, 100)}...`;
      } else {
        resultMessage += ` - Ошибка: Не удалось выполнить экспорт`;
      }
      
      // Добавляем информацию о попытках
      if (currentAttempt > 1) {
        resultMessage += ` (использовано ${currentAttempt - 1} попыток)`;
      }
      
      // Добавляем результат в список
      finishCurrentItemProcessing(resultMessage);
      
      // Переходим к следующему элементу, если он существует
      moveToNextItem();
      
    } catch (error) {
      console.error('Error during auto export:', error);
      finishCurrentItemProcessing(`ID ${selectedItem.Id} - Ошибка экспорта: ${error.message}`);
      moveToNextItem();
    } finally {
      setIsExporting(false);
    }
  };

  // Добавление результата обработки текущего элемента
  const finishCurrentItemProcessing = (result: string): void => {
    setProcessingResults(prev => [...prev, result]);
  };

  // Переход к следующему элементу для автоматической обработки с паузой и отправкой email
  const moveToNextItem = async (): Promise<void> => {
    // Добавляем небольшую паузу перед переходом к следующему элементу
    await new Promise(resolve => setTimeout(resolve, 2000)); // 2 секунды пауза
    
    // Фильтруем элементы с непустым PathForSRSFile
    const validItems = items.filter(item => item.PathForSRSFile && item.PathForSRSFile.trim() !== '');
    
    // Вычисляем индекс следующего элемента
    const nextIndex = currentAutoIndex + 1;
    console.log(`Проверка следующего элемента: ${nextIndex} из ${validItems.length}`);
    
    if (nextIndex < validItems.length) {
      // Переходим к следующему элементу
      setCurrentAutoIndex(nextIndex);
      
      // Выбираем следующий элемент по ключу
      const nextItem = validItems[nextIndex];
      
      // Сбрасываем предыдущие данные для чистой загрузки нового элемента
      setFilteredStaffRecords([]);
      setStaffRecordsGroups([]);
      setLastSearchResult(null);
      
      // Небольшая пауза перед выбором нового элемента
      await new Promise(resolve => setTimeout(resolve, 1000)); // 1 секунда пауза
      
      // Выбираем элемент и устанавливаем selectedItem
      selection.setItems(items);
      selection.setKeySelected(nextItem.key, true, false);
      
      // Еще одна пауза перед установкой selectedItem
      await new Promise(resolve => setTimeout(resolve, 500)); // 0.5 секунд пауза
      
      // Устанавливаем selectedItem, что запустит обработку через useEffect
      setSelectedItem(nextItem);
      
      console.log(`Переходим к элементу ${nextIndex + 1} из ${validItems.length}, ID: ${nextItem.Id}`);
    } else {
      // Завершаем автоматическую обработку
      setIsAutoProcessing(false);
      setCurrentAutoIndex(-1);
      console.log('Автоматическая обработка завершена');
      
      // Формируем результаты для отображения и email
      const title = 'Автоматическая обработка завершена';
      const message = `Обработано элементов: ${validItems.length}\n\nРезультаты:\n${processingResults.join('\n')}`;
      
      // Показываем диалог с результатами
      setDialogTitle(title);
      setDialogMessage(message);
      setIsErrorDialog(false);
      setShowDialog(true);
      
      // Отправляем уведомление по email, если указан адрес
      // Проверяем, есть ли Email у любого из элементов
      const itemsWithEmail = validItems.filter(item => item.email && item.email.trim() !== '');
      
      if (itemsWithEmail.length > 0) {
        // Для каждого уникального Email отправляем уведомление
        const uniqueEmails = new Set<string>();
        itemsWithEmail.forEach(item => {
          if (item.email && item.email.trim().length > 0) {
            uniqueEmails.add(item.email.trim());
          }
        });
        
        const emailAddresses = Array.from(uniqueEmails);
        if (emailAddresses.length > 0) {
          try {
            console.log(`Отправка уведомления на адреса: ${emailAddresses.join(', ')}`);
            const emailBody = formatEmailBody(title, message);
            const emailSent = await sharePointService.sendEmail(emailAddresses, title, emailBody);
            
            if (emailSent) {
              console.log(`Уведомление успешно отправлено на адреса: ${emailAddresses.join(', ')}`);
            } else {
              console.error(`Не удалось отправить уведомление на адреса: ${emailAddresses.join(', ')}`);
            }
          } catch (error) {
            console.error(`Ошибка при отправке уведомления:`, error);
          }
        }
      } else {
        console.log('Нет адресов электронной почты для отправки уведомлений');
      }
    }
  };
  
  // Функция для обработки фильтрации StaffRecords
  const handleFilterStaffRecords = async (selectedExportItem: IExportToSRSItem): Promise<void> => {
    setIsLoadingStaffRecords(true);
    setDebugInfo(null);
    
    try {
      // Загружаем данные StaffRecords, если они еще не загружены
      let records = staffRecords;
      if (records.length === 0) {
        records = await sharePointService.loadStaffRecords();
        setStaffRecords(records);
      }
      
      if (records.length === 0) {
        setFilteredStaffRecords([]);
        setStaffRecordsGroups([]);
        setDebugInfo("Не удалось загрузить данные StaffRecords");
        
        // Если это автоматическая обработка, переходим к следующему элементу
        if (isAutoProcessing) {
          finishCurrentItemProcessing(`ID ${selectedExportItem.Id} - Не удалось загрузить данные StaffRecords`);
          moveToNextItem();
        }
        return;
      }
      
      // Фильтруем записи
      const { filtered, debugInfo } = sharePointService.filterStaffRecords(records, selectedExportItem);
      
      // Создаем группы и сортируем результаты по дате
      const { sortedRecords, groups } = dataUtils.createGroupsFromRecords(filtered);
      
      setFilteredStaffRecords(sortedRecords);
      setStaffRecordsGroups(groups);
      setDebugInfo(debugInfo);
      
      // ВАЖНОЕ ИЗМЕНЕНИЕ: Если записей нет и идет автоматическая обработка, 
      // переходим к следующему элементу
      if (isAutoProcessing && (sortedRecords.length === 0 || groups.length === 0)) {
        finishCurrentItemProcessing(`ID ${selectedExportItem.Id} - Нет соответствующих записей StaffRecords`);
        moveToNextItem();
      }
      
    } catch (error) {
      console.error('Error filtering records:', error);
      setError(`Ошибка при фильтрации записей: ${error.message}`);
      setFilteredStaffRecords([]);
      setStaffRecordsGroups([]);
      
      // Если это автоматическая обработка, переходим к следующему элементу
      if (isAutoProcessing) {
        finishCurrentItemProcessing(`ID ${selectedExportItem.Id} - Ошибка фильтрации: ${error.message}`);
        moveToNextItem();
      }
    } finally {
      setIsLoadingStaffRecords(false);
    }
  };
  
  // Функция для обработки нажатия кнопки Export (принимает groupKey - ключ группы, из которой нажали на экспорт)
  const handleExport = async (groupKey?: string): Promise<void> => {
    if (!selectedItem) {
      setDialogTitle('Ошибка экспорта');
      setDialogMessage('Не выбрана запись для экспорта. Выберите запись из верхней таблицы.');
      setIsErrorDialog(true);
      setShowDialog(true);
      return;
    }

    try {
      setIsExporting(true);
      
      // Получаем путь к файлу из выбранной записи
      const filePath = selectedItem.PathForSRSFile;
      
      // Определяем дату для поиска - из заголовка группы, если указан
      let groupDate: Date | undefined;
      
      if (groupKey) {
        // Пытаемся использовать ключ группы как дату
        try {
          groupDate = new Date(groupKey);
          console.log(`Используем дату группы для экспорта: ${groupKey} (${groupDate.toISOString()})`);
        } catch (e) {
          console.error('Не удалось преобразовать ключ группы в дату:', groupKey, e);
        }
      }
      
      // Проверяем существование файла, передавая дату группы
      const checkResult = await excelService.checkExcelFile(filePath, selectedItem, groupDate);
      
      // Сохраняем результат поиска
      setLastSearchResult(checkResult);
      
      if (checkResult.success) {
        // Определяем заголовок на основе результата поиска строки
        let title = 'Файл найден';
        if (checkResult.rowFound) {
          title = 'Успешно';
        } else {
          title = 'Строка не найдена';
        }
        
        setDialogTitle(title);
        setDialogMessage(checkResult.message);
        setIsErrorDialog(false);
      } else {
        // Файл не найден или произошла ошибка
        setDialogTitle('Ошибка');
        setDialogMessage(checkResult.message);
        setIsErrorDialog(true);
      }
      
      setShowDialog(true);
    } catch (error) {
      console.error('Error during export:', error);
      setDialogTitle('Ошибка экспорта');
      setDialogMessage(`Произошла ошибка при экспорте: ${error.message}`);
      setIsErrorDialog(true);
      setShowDialog(true);
    } finally {
      setIsExporting(false);
    }
  };
  
  // Рендер заголовков групп с кнопкой экспорта
  const onRenderGroupHeader = (props?: IGroupHeaderProps): JSX.Element | null => {
    if (!props) return null;
    
    // Извлекаем key группы (это датa в формате ISO)
    const groupKey = props.group?.key as string;
    
    return (
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', backgroundColor: '#f0f0f0', padding: '10px' }}>
        <GroupHeader 
          {...props} 
          styles={{ 
            root: { 
              backgroundColor: 'transparent', 
              fontWeight: 'bold',
              padding: '0',
              flex: '1'
            },
            title: {
              fontSize: '14px'
            }
          }} 
        />
        <PrimaryButton
          text={isExporting ? "Экспорт..." : "Экспортировать"}
          onClick={() => handleExport(groupKey)} // Передаем ключ группы в метод экспорта
          disabled={isExportDisabled() || isAutoProcessing}
          styles={{ root: { minWidth: '120px' } }}
        />
      </div>
    );
  };

  // Закрытие диалога
  const closeDialog = (): void => {
    setShowDialog(false);
  };

  // Проверка наличия данных в PathForSRSFile у выбранной записи
  const isExportDisabled = (): boolean => {
    return !selectedItem || isExporting || !selectedItem.PathForSRSFile || selectedItem.PathForSRSFile.trim() === '';
  };

  return (
    <div className={styles.kpfasrs}>
      <div className={styles.container}>
        <div className={styles.row}>
          <div className={styles.column}>
            <h2>ExportToSRS Data</h2>
            
            {/* Кнопка для загрузки данных в первую таблицу */}
            <div className={styles.buttonContainer}>
              <PrimaryButton
                text={isLoading ? "Загрузка..." : "Загрузить данные из ExportToSRS"}
                onClick={handleLoadExportToSRSItems}
                disabled={isLoading || isAutoProcessing}
                className={styles.loadButton}
              />
            </div>
            
            {/* Индикатор автоматической обработки */}
            {isAutoProcessing && (
              <MessageBar
                messageBarType={MessageBarType.info}
                isMultiline={false}
                className={styles.infoMessage}
              >
                Выполняется автоматическая обработка... Элемент {currentAutoIndex + 1} из {items.filter(item => item.PathForSRSFile && item.PathForSRSFile.trim() !== '').length}
              </MessageBar>
            )}
            
            {error && (
              <MessageBar
                messageBarType={MessageBarType.error}
                isMultiline={true}
                dismissButtonAriaLabel="Close"
                className={styles.errorMessage}
              >
                {error}
              </MessageBar>
            )}
            
            {isLoading ? (
              <Spinner label={`Loading data from SharePoint...`} size={SpinnerSize.medium} />
            ) : (
              <>
                {items.length === 0 && !error ? (
                  <MessageBar
                    messageBarType={MessageBarType.warning}
                    className={styles.warningMessage}
                  >
                    Нажмите кнопку "Загрузить данные из ExportToSRS" для отображения записей.
                  </MessageBar>
                ) : (
                  <div className={styles.tableContainer}>
                    <h3>ExportToSRS Records</h3>
                    <DetailsList
                      items={items}
                      columns={columns}
                      selection={selection}
                      selectionMode={SelectionMode.single}
                      setKey="id"
                      className={styles.detailsListWrapper}
                    />
                    
                    <h3 className={styles.secondTableHeader}>Related StaffRecords</h3>
                    {isLoadingStaffRecords ? (
                      <Spinner label="Loading related records..." size={SpinnerSize.medium} />
                    ) : (
                      <>
                        {selectedItem ? (
                          <>
                            {lastSearchResult && (
                              <MessageBar
                                messageBarType={lastSearchResult.rowFound ? MessageBarType.success : MessageBarType.warning}
                                isMultiline={true}
                                className={styles.infoMessage}
                                dismissButtonAriaLabel="Close"
                              >
                                <h4>Результат поиска в Excel:</h4>
                                <div>
                                  {lastSearchResult.rowFound ? 
                                    `Строка найдена в позиции ${lastSearchResult.rowNumber}` : 
                                    `Строка не найдена. Проверьте формат даты в Excel.`
                                  }
                                </div>
                              </MessageBar>
                            )}
                            
                            {/* Результаты автоматической обработки */}
                            {isAutoProcessing && processingResults.length > 0 && (
                              <MessageBar
                                messageBarType={MessageBarType.info}
                                isMultiline={true}
                                className={styles.debugMessage}
                              >
                                <h4>Результаты обработки:</h4>
                                <pre className={styles.debugPre}>
                                  {processingResults.join('\n')}
                                </pre>
                              </MessageBar>
                            )}
                            
                            {debugInfo && (
                              <MessageBar
                                messageBarType={MessageBarType.info}
                                isMultiline={true}
                                className={styles.debugMessage}
                              >
                                <h4>Отладочная информация:</h4>
                                <pre className={styles.debugPre}>
                                  {debugInfo}
                                </pre>
                              </MessageBar>
                            )}
                            
                            {filteredStaffRecords.length === 0 ? (
                              <MessageBar
                                messageBarType={MessageBarType.info}
                                className={styles.infoMessage}
                              >
                                No matching StaffRecords found for the selected criteria.
                              </MessageBar>
                            ) : (
                              <DetailsList
                                items={filteredStaffRecords}
                                columns={staffRecordsColumns}
                                groups={staffRecordsGroups}
                                groupProps={{
                                  showEmptyGroups: true,
                                  onRenderHeader: onRenderGroupHeader
                                }}
                                selectionMode={SelectionMode.none}
                                setKey="staffRecordsId"
                                className={styles.detailsListWrapper}
                              />
                            )}
                          </>
                        ) : (
                          <MessageBar
                            messageBarType={MessageBarType.info}
                            className={styles.infoMessage}
                          >
                            Select an item from the ExportToSRS table to see related StaffRecords.
                          </MessageBar>
                        )}
                      </>
                    )}
                  </div>
                )}
              </>
            )}
          </div>
        </div>
      </div>
      
      {/* Диалог для результатов экспорта */}
      <Dialog
        hidden={!showDialog}
        onDismiss={closeDialog}
        dialogContentProps={{
          type: isErrorDialog ? DialogType.largeHeader : DialogType.normal,
          title: dialogTitle,
          subText: dialogMessage,
          styles: isErrorDialog 
            ? { title: { color: '#a4262c' } } 
            : { title: { color: lastSearchResult?.rowFound ? '#107c10' : '#f3901d' } }
        }}
        modalProps={{
          isBlocking: false,
          styles: { main: { maxWidth: 600 } }
        }}
      >
        <DialogFooter>
          <DefaultButton onClick={closeDialog} text="Закрыть" />
        </DialogFooter>
      </Dialog>
    </div>
  );
};

export default Kpfasrs;