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
  Dialog,
  DialogType,
  DialogFooter,
  DefaultButton
} from '@fluentui/react';
import { SharePointService } from '../services/SharePointService';
import { ExcelService, IExportToSRSItem, IStaffRecordsItem, IFileCheckResult } from '../services/ExcelService';
import { DataUtils } from '../services/DataUtils';

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
  
  // Состояния для экспорта и диалогов
  const [isExporting, setIsExporting] = useState<boolean>(false);
  const [showDialog, setShowDialog] = useState<boolean>(false);
  const [dialogMessage, setDialogMessage] = useState<string>('');
  const [dialogTitle, setDialogTitle] = useState<string>('');
  const [isErrorDialog, setIsErrorDialog] = useState<boolean>(false);
  const [lastSearchResult, setLastSearchResult] = useState<IFileCheckResult | null>(null);
  
  // Общие состояния
  const [error, setError] = useState<string | null>(null);
  const [debugInfo, setDebugInfo] = useState<string | null>(null);

  // Инициализация сервисов
  const sharePointService = React.useMemo(() => new SharePointService(props.context), [props.context]);
  const excelService = React.useMemo(() => new ExcelService(props.context), [props.context]);
  const dataUtils = React.useMemo(() => new DataUtils(excelService), [excelService]);

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
    { key: 'pathForSRSFile', name: 'Path For SRS File', fieldName: 'PathForSRSFile', minWidth: 200, maxWidth: 300 }
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
  
  // Загрузка данных из списка ExportToSRS при инициализации
  useEffect(() => {
    const loadExportToSRSItems = async (): Promise<void> => {
      try {
        setIsLoading(true);
        setError(null);
        
        const loadedItems = await sharePointService.getExportToSRSItems();
        setItems(loadedItems);
      } catch (error) {
        console.error('Error in loadExportToSRSItems:', error);
        setError(`Ошибка: ${error.message}`);
      } finally {
        setIsLoading(false);
      }
    };

    loadExportToSRSItems();
  }, [sharePointService]);
  
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
        return;
      }
      
      // Фильтруем записи
      const { filtered, debugInfo } = sharePointService.filterStaffRecords(records, selectedExportItem);
      
      // Создаем группы и сортируем результаты по дате
      const { sortedRecords, groups } = dataUtils.createGroupsFromRecords(filtered);
      
      setFilteredStaffRecords(sortedRecords);
      setStaffRecordsGroups(groups);
      setDebugInfo(debugInfo);
      
    } catch (error) {
      console.error('Error filtering records:', error);
      setError(`Ошибка при фильтрации записей: ${error.message}`);
      setFilteredStaffRecords([]);
      setStaffRecordsGroups([]);
    } finally {
      setIsLoadingStaffRecords(false);
    }
  };
  
  // Функция для обработки нажатия кнопки Export (теперь принимает groupKey - ключ группы, из которой нажали на экспорт)
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
          disabled={isExportDisabled()}
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