import * as React from 'react';
import { useState, useEffect } from 'react';
import styles from './Kpfasrs.module.scss';
import { IKpfasrsProps } from './IKpfasrsProps';
import { 
  DetailsList, 
  SelectionMode, 
  IColumn,
  PrimaryButton,
  Dialog,
  DialogType,
  DialogFooter,
  DefaultButton,
  MessageBar,
  MessageBarType
} from '@fluentui/react';
import { SPService } from '../services/SPService';

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

const Kpfasrs: React.FC<IKpfasrsProps> = (props) => {
  const [items, setItems] = useState<IExportToSRSItem[]>([]);
  const [isLoading, setIsLoading] = useState<boolean>(true);
  const [showDialog, setShowDialog] = useState<boolean>(false);
  const [dialogMessage, setDialogMessage] = useState<string>('');

  // Определение колонок для DetailsList
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

  // Состояние для хранения ошибок
  const [error, setError] = useState<string | null>(null);

  // Загрузка данных из списка SharePoint
  useEffect(() => {
    const getListItems = async (): Promise<void> => {
      try {
        setIsLoading(true);
        setError(null);
        
        // Используем наш сервис для получения данных
        const spService = new SPService(props.context);
        
        console.log('Fetching data from ExportToSRS list at:', kpfaDataUrl);
        
        // Указываем только нужные поля, включая ID для связанных элементов
        const select = "Id,Title,StaffMemberId,Date1,Date2,ManagerId,StaffGroupId,Condition,GroupMemberId";
        
        const listItems = await spService.getListItems(
          kpfaDataUrl,
          'ExportToSRS',
          select,
          undefined,
          "Id"
        );
        
        console.log('Data loaded successfully:', listItems.length, 'items');
        console.log('Sample data item:', listItems.length > 0 ? listItems[0] : 'No items');
        
        setItems(listItems);
        
      } catch (error) {
        console.error('Error in getListItems:', error);
        setError(`Ошибка при загрузке данных: ${error.message}`);
        setDialogMessage(`Ошибка при загрузке данных: ${error.message}`);
        setShowDialog(true);
      } finally {
        setIsLoading(false);
      }
    };

    getListItems();
  }, [props.context]);

  // Функция для отправки данных в Office Script
  const exportToOfficeScript = async (): Promise<void> => {
    try {
      setIsLoading(true);
      
      // В рабочей версии здесь должна быть логика отправки данных в Office Script
      // Пример подготовки данных с ID для полей Lookup:
      const dataToExport = items.map(item => ({
        Id: item.Id,
        Title: item.Title ? 'Yes' : 'No',
        StaffMember: item.StaffMemberId,
        Date1: item.Date1,
        Date2: item.Date2,
        Manager: item.ManagerId,
        StaffGroup: item.StaffGroupId, 
        Condition: item.Condition,
        GroupMember: item.GroupMemberId
      }));
      
      console.log('Data prepared for export:', dataToExport);
      
      // Так как это демо без реального API Office Script,
      // просто показываем диалог об успешном экспорте
      setDialogMessage('Данные готовы для экспорта в Excel через Office Script!');
      setShowDialog(true);
      
    } catch (error) {
      console.error('Error exporting to Office Script:', error);
      setDialogMessage('Ошибка при экспорте: ' + error.message);
      setShowDialog(true);
    } finally {
      setIsLoading(false);
    }
  };

  // Закрытие диалога
  const closeDialog = (): void => {
    setShowDialog(false);
  };

  return (
    <div className={styles.kpfasrs}>
      <div className={styles.container}>
        <div className={styles.row}>
          <div className={styles.column}>
            <h2>ExportToSRS Data</h2>
            
            <div className={styles.buttonContainer}>
              <PrimaryButton 
                text="Export to Excel via Office Script" 
                onClick={exportToOfficeScript} 
                disabled={isLoading || items.length === 0}
                className={styles.exportButton}
              />
            </div>
            
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
              <p>Loading data from {kpfaDataUrl}...</p>
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
                  <DetailsList
                    items={items}
                    columns={columns}
                    selectionMode={SelectionMode.none}
                    setKey="id"
                  />
                )}
              </>
            )}
          </div>
        </div>
      </div>
      
      <Dialog
        hidden={!showDialog}
        onDismiss={closeDialog}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'Export Status',
          subText: dialogMessage
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