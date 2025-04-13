import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IExportToSRSItem, IStaffRecordsItem } from './ExcelService';

// URL вашего сайта SharePoint с данными
const kpfaDataUrl: string = "https://kpfaie.sharepoint.com/sites/KPFAData";

export class SharePointService {
  private context: WebPartContext;
  
  constructor(context: WebPartContext) {
    this.context = context;
  }
  
  // Загрузка данных из списка ExportToSRS
  public async getExportToSRSItems(): Promise<IExportToSRSItem[]> {
    try {
      console.log('Fetching data from ExportToSRS list at:', kpfaDataUrl);
      
      const endpoint = `${kpfaDataUrl}/_api/web/lists/getbytitle('ExportToSRS')/items`;
      const select = "Id,Title,StaffMemberId,Date1,Date2,ManagerId,StaffGroupId,Condition,GroupMemberId,PathForSRSFile";
      const queryUrl = `${endpoint}?$select=${select}`;
      
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        queryUrl,
        SPHttpClient.configurations.v1
      );

      if (response.ok) {
        const results = await response.json();
        console.log('Data loaded successfully:', results.value.length, 'items');
        console.log('Sample data item:', results.value.length > 0 ? results.value[0] : 'No items');
        return results.value;
      } else {
        const errorText = await response.text();
        console.error('Error fetching ExportToSRS items:', response.status, errorText);
        throw new Error(`Ошибка при загрузке данных из ExportToSRS: ${response.status} ${response.statusText}`);
      }
    } catch (error) {
      console.error('Error in getExportToSRSItems:', error);
      throw error;
    }
  }

  // Функция для загрузки данных из StaffRecords
  public async loadStaffRecords(): Promise<IStaffRecordsItem[]> {
    try {
      console.log('Loading StaffRecords from:', kpfaDataUrl);
      
      // Массив для всех записей
      let allRecords: IStaffRecordsItem[] = [];
      
      // URL для первого запроса
      let endpoint = `${kpfaDataUrl}/_api/web/lists/getbytitle('StaffRecords')/items`;
      const select = "Id,Title,Date,StaffMemberId,StaffMember/Id,StaffMember/Title,ManagerId,StaffGroupId," +
                    "Checked,ExportResult,ShiftDate1,ShiftDate2,TimeForLunch,Contract,TypeOfLeaveId,TypeOfLeave/Id," +
                    "TypeOfLeave/Title,LeaveTime,LeaveNote,LunchNote,TotalHoursNote,ReliefHours";
      const expand = "StaffMember,TypeOfLeave";
      let queryUrl = `${endpoint}?$select=${select}&$expand=${expand}&$top=5000`;
      
      let nextLink: string | null = queryUrl;
      let pageCount = 1;
      
      // Цикл для обработки всех страниц
      while (nextLink) {
        console.log(`Loading StaffRecords page ${pageCount}...`);
        
        const response: SPHttpClientResponse = await this.context.spHttpClient.get(
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
      
      return allRecords;
    } catch (error) {
      console.error('Error loading StaffRecords:', error);
      throw error;
    }
  }

  // Функция для фильтрации StaffRecords на основе выбранной строки ExportToSRS
  public filterStaffRecords(
    records: IStaffRecordsItem[], 
    selectedExportItem: IExportToSRSItem
  ): { 
    filtered: IStaffRecordsItem[], 
    debugInfo: string 
  } {
    console.log('Filtering staff records for:', selectedExportItem);
    
    // Подробная информация о выбранной записи для отладки
    console.log('Selected record details:', {
      Id: selectedExportItem.Id,
      Date1: selectedExportItem.Date1,
      Date2: selectedExportItem.Date2,
      ManagerId: selectedExportItem.ManagerId,
      StaffGroupId: selectedExportItem.StaffGroupId,
      StaffMemberId: selectedExportItem.StaffMemberId,
      PathForSRSFile: selectedExportItem.PathForSRSFile
    });
    
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
    debugLines.push(`Path For SRS File: ${selectedExportItem.PathForSRSFile || 'Not specified'}`);
    
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
      debugLines.push(`- TimeForLunch: ${filtered[0].TimeForLunch}`);
      debugLines.push(`- LeaveTime: ${filtered[0].LeaveTime}`);
      debugLines.push(`- ReliefHours: ${filtered[0].ReliefHours}`);
      debugLines.push(`- LeaveNote: ${filtered[0].LeaveNote ? filtered[0].LeaveNote.substring(0, 30) + '...' : ''}`);
      debugLines.push(`- Checked: ${filtered[0].Checked}`);
      debugLines.push(`- ExportResult: ${filtered[0].ExportResult}`);
    }
    
    console.log(debugLines.join('\n'));
    console.log(`Found ${filtered.length} matching StaffRecords with all conditions`);
    
    return {
      filtered,
      debugInfo: debugLines.join('\n')
    };
  }
}