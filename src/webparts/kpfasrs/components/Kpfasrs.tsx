import * as React from 'react';
import { useState, useEffect } from 'react';
import styles from './Kpfasrs.module.scss';
import { IKpfasrsProps } from './IKpfasrsProps';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
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
  IGroupHeaderProps
} from '@fluentui/react';

// URL вашего сайта SharePoint
const kpfaDataUrl: string = "https://kpfaie.sharepoint.com/sites/KPFAData";

// Интерфейс для данных из списка ExportToSRS
interface IExportToSRSItem {
  Id: number;
  Title: boolean;
  StaffMemberId: number;
  Date1: string;
  Date2: string;
  ManagerId: number;
  StaffGroupId: number;
  Condition: number;
  GroupMemberId: number;
}

// Интерфейс для данных из списка StaffRecords
interface IStaffRecordsItem {
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
  // Новые поля
  ShiftDate1: string;
  ShiftDate2: string;
  TimeForLunch: number;  // Исправлено с LunchTime на TimeForLunch
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

const Kpfasrs: React.FC<IKpfasrsProps> = (props) => {
  // Состояния для ExportToSRS
  const [items, setItems] = useState<IExportToSRSItem[]>([]);
  const [isLoading, setIsLoading] = useState<boolean>(true);
  
  // Состояния для StaffRecords
  const [staffRecords, setStaffRecords] = useState<IStaffRecordsItem[]>([]);
  const [filteredStaffRecords, setFilteredStaffRecords] = useState<IStaffRecordsItem[]>([]);
  const [staffRecordsGroups, setStaffRecordsGroups] = useState<IGroup[]>([]);
  const [isLoadingStaffRecords, setIsLoadingStaffRecords] = useState<boolean>(false);
  const [selectedItem, setSelectedItem] = useState<IExportToSRSItem | null>(null);
  
  // Общие состояния
  const [error, setError] = useState<string | null>(null);
  const [debugInfo, setDebugInfo] = useState<string | null>(null);

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
    { key: 'groupMember', name: 'Group Member', fieldName: 'GroupMemberId', minWidth: 100, maxWidth: 150 }
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
    { key: 'timeForLunch', name: 'Time For Lunch', fieldName: 'TimeForLunch', minWidth: 100, maxWidth: 120 },  // Исправлено
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
      const selectedItems = selection.getSelection() as IExportToSRSItem[];
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

  // Загрузка данных из списка ExportToSRS
  useEffect(() => {
    const getExportToSRSItems = async (): Promise<void> => {
      try {
        setIsLoading(true);
        setError(null);
        
        console.log('Fetching data from ExportToSRS list at:', kpfaDataUrl);
        
        const endpoint = `${kpfaDataUrl}/_api/web/lists/getbytitle('ExportToSRS')/items`;
        const select = "Id,Title,StaffMemberId,Date1,Date2,ManagerId,StaffGroupId,Condition,GroupMemberId";
        const queryUrl = `${endpoint}?$select=${select}`;
        
        const response: SPHttpClientResponse = await props.context.spHttpClient.get(
          queryUrl,
          SPHttpClient.configurations.v1
        );

        if (response.ok) {
          const results = await response.json();
          console.log('Data loaded successfully:', results.value.length, 'items');
          console.log('Sample data item:', results.value.length > 0 ? results.value[0] : 'No items');
          setItems(results.value);
        } else {
          const errorText = await response.text();
          console.error('Error fetching ExportToSRS items:', response.status, errorText);
          setError(`Ошибка при загрузке данных из ExportToSRS: ${response.status} ${response.statusText}`);
        }
      } catch (error) {
        console.error('Error in getExportToSRSItems:', error);
        setError(`Ошибка: ${error.message}`);
      } finally {
        setIsLoading(false);
      }
    };

    getExportToSRSItems();
  }, [props.context]);

  // Функция для загрузки данных из StaffRecords
  const loadStaffRecords = async (): Promise<IStaffRecordsItem[]> => {
    try {
      setIsLoadingStaffRecords(true);
      
      console.log('Loading StaffRecords from:', kpfaDataUrl);
      
      // Массив для всех записей
      let allRecords: IStaffRecordsItem[] = [];
      
      // URL для первого запроса
      let endpoint = `${kpfaDataUrl}/_api/web/lists/getbytitle('StaffRecords')/items`;
      const select = "Id,Title,Date,StaffMemberId,StaffMember/Id,StaffMember/Title,ManagerId,StaffGroupId," +
                    "Checked,ExportResult,ShiftDate1,ShiftDate2,TimeForLunch,Contract,TypeOfLeaveId,TypeOfLeave/Id," +  // Исправлено
                    "TypeOfLeave/Title,LeaveTime,LeaveNote,LunchNote,TotalHoursNote,ReliefHours";
      const expand = "StaffMember,TypeOfLeave";
      let queryUrl = `${endpoint}?$select=${select}&$expand=${expand}&$top=5000`; // Увеличиваем лимит до 5000 (максимум для SharePoint)
      
      let nextLink: string | null = queryUrl;
      let pageCount = 1;
      
      // Цикл для обработки всех страниц
      while (nextLink) {
        console.log(`Loading StaffRecords page ${pageCount}...`);
        
        const response: SPHttpClientResponse = await props.context.spHttpClient.get(
          nextLink,
          SPHttpClient.configurations.v1
        );
        
        if (!response.ok) {
          const errorText = await response.text();
          console.error(`Error fetching StaffRecords page ${pageCount}:`, response.status, errorText);
          throw new Error(`Ошибка при загрузке данных StaffRecords: ${response.status} ${response.statusText}`);
        }
        
        const results = await response.json();
        const records = results.value;
        console.log(`Loaded ${records.length} StaffRecords on page ${pageCount}`);
        
        // Добавляем записи к общему массиву
        allRecords = [...allRecords, ...records];
        
        // Проверяем, есть ли еще страницы
        nextLink = results["@odata.nextLink"] || null;
        pageCount++;
        
        // Если слишком много страниц, прерываем цикл для безопасности
        if (pageCount > 20) {
          console.warn('Too many pages of StaffRecords, stopping at 20 pages.');
          break;
        }
      }
      
      console.log(`Total StaffRecords loaded: ${allRecords.length}`);
      
      // Выводим пример записи для отладки
      if (allRecords.length > 0) {
        console.log('StaffRecords sample (first record):', allRecords[0]);
      }
      
      // Проверяем, есть ли записи со StaffGroupId = 54
      const recordsWithStaffGroup54 = allRecords.filter(r => r.StaffGroupId === 54);
      console.log(`StaffRecords with StaffGroupId = 54: ${recordsWithStaffGroup54.length}`);
      
      // Находим уникальные значения StaffGroupId
      const groupIdMap: { [key: string]: boolean } = {};
      allRecords.forEach(r => {
        groupIdMap[String(r.StaffGroupId)] = true;
      });
      const uniqueGroupIds = Object.keys(groupIdMap).map(key => isNaN(Number(key)) ? key : Number(key));
      console.log('Unique StaffGroupId values in all StaffRecords:', uniqueGroupIds);
      
      setStaffRecords(allRecords);
      return allRecords;
    } catch (error) {
      console.error('Error loading StaffRecords:', error);
      setError(`Ошибка при загрузке StaffRecords: ${error.message}`);
      return [];
    } finally {
      setIsLoadingStaffRecords(false);
    }
  };

  // Функция для создания групп на основе дат
// Функция для проверки, является ли смена "пустой" (время 00:00)
const isEmptyShift = (record: IStaffRecordsItem): boolean => {
  const shiftDate1 = record.ShiftDate1 || '';
  const shiftDate2 = record.ShiftDate2 || '';
  
  // Используем indexOf вместо endsWith для лучшей совместимости
  const isShiftDate1Empty = shiftDate1.indexOf('T00:00:00Z') === shiftDate1.length - 10 || 
                          shiftDate1.indexOf('T00:00:00') === shiftDate1.length - 9;
  
  const isShiftDate2Empty = shiftDate2.indexOf('T00:00:00Z') === shiftDate2.length - 10 || 
                          shiftDate2.indexOf('T00:00:00') === shiftDate2.length - 9;
  
  return isShiftDate1Empty && isShiftDate2Empty;
};

// Функция для создания групп на основе дат
const createGroupsFromRecords = (records: IStaffRecordsItem[]): void => {
  if (!records || records.length === 0) {
    setStaffRecordsGroups([]);
    return;
  }
  
  // Сортируем записи по дате для группировки
  const recordsByDate = [...records].sort((a, b) => {
    const dateA = new Date(a.Date);
    const dateB = new Date(b.Date);
    return dateA.getTime() - dateB.getTime();
  });
  
  // Группируем записи по дате
  const recordGroups: { [key: string]: IStaffRecordsItem[] } = {};
  recordsByDate.forEach(record => {
    // Извлекаем только дату без времени
    const datePart = record.Date.split('T')[0];
    
    if (!recordGroups[datePart]) {
      recordGroups[datePart] = [];
    }
    
    recordGroups[datePart].push(record);
  });
  
  // Сортируем записи внутри каждой группы:
  // - Сначала по ShiftDate1
  // - Пустые смены в конце
  Object.keys(recordGroups).forEach(date => {
    recordGroups[date].sort((a, b) => {
      // Проверяем, является ли какая-либо из смен "пустой"
      const aEmpty = isEmptyShift(a);
      const bEmpty = isEmptyShift(b);
      
      // Если одна пустая, а другая нет, пустая идет в конец
      if (aEmpty && !bEmpty) return 1;
      if (!aEmpty && bEmpty) return -1;
      
      // Если обе пустые или обе не пустые, сортируем по ShiftDate1
      const timeA = new Date(a.ShiftDate1 || a.Date).getTime();
      const timeB = new Date(b.ShiftDate1 || b.Date).getTime();
      return timeA - timeB;
    });
  });
  
  // Готовим отсортированный массив всех записей
  const sortedRecords: IStaffRecordsItem[] = [];
  const dates = Object.keys(recordGroups).sort(); // Сортируем даты
  
  dates.forEach(date => {
    sortedRecords.push(...recordGroups[date]);
  });
  
  // Сохраняем отсортированные записи
  setFilteredStaffRecords(sortedRecords);
  
  // Создаем группы для DetailsList
  const groups: IGroup[] = [];
  let startIndex = 0;
  
  dates.forEach(date => {
    const count = recordGroups[date].length;
    
    groups.push({
      key: date,
      name: formatDate(date),
      startIndex,
      count,
      level: 0,
      isCollapsed: false
    });
    
    startIndex += count;
  });
  
  console.log('Created groups with custom sorting:', groups);
  setStaffRecordsGroups(groups);
};
  
  // Функция для форматирования даты
  const formatDate = (dateString: string): string => {
    const options: Intl.DateTimeFormatOptions = { 
      weekday: 'long', 
      year: 'numeric', 
      month: 'long', 
      day: 'numeric' 
    };
    
    try {
      const date = new Date(dateString);
      return date.toLocaleDateString('ru-RU', options);
    } catch (e) {
      console.error('Error formatting date:', dateString, e);
      return dateString;
    }
  };

  // Функция для фильтрации StaffRecords на основе выбранной строки ExportToSRS
  const handleFilterStaffRecords = async (selectedExportItem: IExportToSRSItem): Promise<void> => {
    console.log('Filtering staff records for:', selectedExportItem);
    
    // Подробная информация о выбранной записи для отладки
    console.log('Selected record details:', {
      Id: selectedExportItem.Id,
      Date1: selectedExportItem.Date1,
      Date2: selectedExportItem.Date2,
      ManagerId: selectedExportItem.ManagerId,
      StaffGroupId: selectedExportItem.StaffGroupId,
      StaffMemberId: selectedExportItem.StaffMemberId
    });
    
    setIsLoadingStaffRecords(true);
    setDebugInfo(null);
    
    try {
      // Загружаем данные StaffRecords, если они еще не загружены
      let records = staffRecords;
      if (records.length === 0) {
        records = await loadStaffRecords();
      }
      
      if (records.length === 0) {
        setFilteredStaffRecords([]);
        setStaffRecordsGroups([]);
        setDebugInfo("Не удалось загрузить данные StaffRecords");
        return;
      }
      
      // Подготавливаем строки с информацией об отладке
      let debugLines: string[] = [];
      
      // Преобразуем строковые даты в объекты Date для сравнения
      let date1: Date;
      let date2: Date;
      
      try {
        date1 = new Date(selectedExportItem.Date1);
        date2 = new Date(selectedExportItem.Date2);
        
        debugLines.push(`Date1 (исходная): ${selectedExportItem.Date1}`);
        debugLines.push(`Date1 (преобразованная): ${date1.toISOString()}`);
        debugLines.push(`Date2 (исходная): ${selectedExportItem.Date2}`);
        debugLines.push(`Date2 (преобразованная): ${date2.toISOString()}`);
      } catch (e) {
        debugLines.push(`Ошибка при преобразовании дат: ${e.message}`);
        date1 = new Date(0); // Минимальная дата
        date2 = new Date(); // Текущая дата
      }
      
      // Находим уникальные значения StaffGroupId в записях
      const groupIdMap: { [key: string]: boolean } = {};
      records.forEach(r => {
        groupIdMap[String(r.StaffGroupId)] = true;
      });
      const uniqueGroupIds = Object.keys(groupIdMap).map(key => isNaN(Number(key)) ? key : Number(key));
      debugLines.push(`Уникальные значения StaffGroupId в StaffRecords: ${uniqueGroupIds.join(', ')}`);
      
      // Проверка каждого условия фильтрации отдельно
      const matchingDate = records.filter(record => {
        try {
          const recordDate = new Date(record.Date);
          return recordDate >= date1 && recordDate <= date2;
        } catch (e) {
          console.error('Error comparing dates for record:', record.Id, e);
          return false;
        }
      });
      
      const matchingManager = records.filter(record => 
        record.ManagerId === selectedExportItem.ManagerId
      );
      
      const matchingGroup = records.filter(record => 
        record.StaffGroupId === selectedExportItem.StaffGroupId
      );
      
      const matchingStaffMember = records.filter(record => 
        record.StaffMemberId === selectedExportItem.StaffMemberId
      );
      
      debugLines.push(`Всего записей StaffRecords: ${records.length}`);
      debugLines.push(`Записей, соответствующих условию по дате: ${matchingDate.length}`);
      debugLines.push(`Записей, соответствующих условию по ManagerId (${selectedExportItem.ManagerId}): ${matchingManager.length}`);
      debugLines.push(`Записей, соответствующих условию по StaffGroupId (${selectedExportItem.StaffGroupId}): ${matchingGroup.length}`);
      debugLines.push(`Записей, соответствующих условию по StaffMemberId (${selectedExportItem.StaffMemberId}): ${matchingStaffMember.length}`);
      
      // Применяем все условия фильтрации
      const filtered = records.filter((record: IStaffRecordsItem) => {
        try {
          const recordDate = new Date(record.Date);
          
          const dateCondition = recordDate >= date1 && recordDate <= date2;
          const managerCondition = record.ManagerId === selectedExportItem.ManagerId;
          const groupCondition = record.StaffGroupId === selectedExportItem.StaffGroupId;
          const staffMemberCondition = record.StaffMemberId === selectedExportItem.StaffMemberId;
          
          return dateCondition && managerCondition && groupCondition && staffMemberCondition;
        } catch (e) {
          console.error('Error filtering record:', record.Id, e);
          return false;
        }
      });
      
      debugLines.push(`Найдено записей, соответствующих всем условиям: ${filtered.length}`);
      
      if (filtered.length > 0) {
        debugLines.push(`Пример записи, соответствующей всем условиям:`);
        debugLines.push(`- Id: ${filtered[0].Id}`);
        debugLines.push(`- Date: ${filtered[0].Date}`);
        debugLines.push(`- ShiftDate1: ${filtered[0].ShiftDate1}`);
        debugLines.push(`- ShiftDate2: ${filtered[0].ShiftDate2}`);
        debugLines.push(`- ManagerId: ${filtered[0].ManagerId}`);
        debugLines.push(`- StaffGroupId: ${filtered[0].StaffGroupId}`);
        debugLines.push(`- StaffMemberId: ${filtered[0].StaffMemberId}`);
        debugLines.push(`- TypeOfLeaveId: ${filtered[0].TypeOfLeaveId}`);
        debugLines.push(`- Contract: ${filtered[0].Contract}`);
        debugLines.push(`- TimeForLunch: ${filtered[0].TimeForLunch}`);  // Исправлено
        debugLines.push(`- LeaveTime: ${filtered[0].LeaveTime}`);
        debugLines.push(`- ReliefHours: ${filtered[0].ReliefHours}`);
        debugLines.push(`- LeaveNote: ${filtered[0].LeaveNote ? filtered[0].LeaveNote.substring(0, 30) + '...' : ''}`);
        debugLines.push(`- Checked: ${filtered[0].Checked}`);
        debugLines.push(`- ExportResult: ${filtered[0].ExportResult}`);
      }
      
      console.log(debugLines.join('\n'));
      setDebugInfo(debugLines.join('\n'));
      
      console.log(`Found ${filtered.length} matching StaffRecords with all conditions`);
      
      // Создаем группы и сортируем результаты по дате
      createGroupsFromRecords(filtered);
      
    } catch (error) {
      console.error('Error filtering records:', error);
      setError(`Ошибка при фильтрации записей: ${error.message}`);
      setFilteredStaffRecords([]);
      setStaffRecordsGroups([]);
    } finally {
      setIsLoadingStaffRecords(false);
    }
  };

  // Кастомный рендер для заголовка группы
  const onRenderGroupHeader = (props?: IGroupHeaderProps): JSX.Element | null => {
    if (!props) return null;
    
    return (
      <GroupHeader 
        {...props} 
        styles={{ 
          root: { 
            backgroundColor: '#f0f0f0', 
            fontWeight: 'bold',
            padding: '10px 0'
          },
          title: {
            fontSize: '14px'
          }
        }} 
      />
    );
  };

  return (
    <div className={styles.kpfasrs}>
      <div className={styles.container}>
        <div className={styles.row}>
          <div className={styles.column}>
            <h2>ExportToSRS Data</h2>
            
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
              <Spinner label={`Loading data from ${kpfaDataUrl}...`} size={SpinnerSize.medium} />
            ) : (
              <>
                {items.length === 0 && !error ? (
                  <MessageBar
                    messageBarType={MessageBarType.warning}
                    className={styles.warningMessage}
                  >
                    No data found in the ExportToSRS list or you don't have permissions to access it.
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
    </div>
  );
};

export default Kpfasrs;