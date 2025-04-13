import { IGroup } from '@fluentui/react';
import { IStaffRecordsItem, ExcelService } from './ExcelService';

export class DataUtils {
  private excelService: ExcelService;
  
  constructor(excelService: ExcelService) {
    this.excelService = excelService;
  }
  
  // Функция для создания групп на основе дат
  public createGroupsFromRecords(records: IStaffRecordsItem[]): { 
    sortedRecords: IStaffRecordsItem[], 
    groups: IGroup[] 
  } {
    if (!records || records.length === 0) {
      return { sortedRecords: [], groups: [] };
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
        const aEmpty = this.excelService.isEmptyShift(a);
        const bEmpty = this.excelService.isEmptyShift(b);
        
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
    
    // Создаем группы для DetailsList
    const groups: IGroup[] = [];
    let startIndex = 0;
    
    dates.forEach(date => {
      const count = recordGroups[date].length;
      
      groups.push({
        key: date,
        name: this.formatDate(date),
        startIndex,
        count,
        level: 0,
        isCollapsed: false
      });
      
      startIndex += count;
    });
    
    console.log('Created groups with custom sorting:', groups);
    
    return {
      sortedRecords,
      groups
    };
  }
  
  // Функция для форматирования даты
  public formatDate(dateString: string): string {
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
  }
}