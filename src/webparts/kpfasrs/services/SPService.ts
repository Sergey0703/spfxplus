import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export class SPService {
  private context: WebPartContext;
  
  constructor(context: WebPartContext) {
    this.context = context;
  }
  
  /**
   * Получает элементы из списка SharePoint
   * @param siteUrl URL сайта SharePoint
   * @param listName Имя списка
   * @param select Поля для выбора (опционально)
   * @param filter Условие фильтрации (опционально)
   * @param orderBy Поле для сортировки (опционально)
   * @returns Массив элементов списка
   */
  public async getListItems(
    siteUrl: string, 
    listName: string, 
    select?: string, 
    filter?: string, 
    orderBy?: string
  ): Promise<any[]> {
    try {
      let endpoint = `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items`;
      
      // Добавляем параметры запроса, если они указаны
      const queryParams: string[] = [];
      
      if (select) {
        queryParams.push(`$select=${select}`);
      }
      
      if (filter) {
        queryParams.push(`$filter=${filter}`);
      }
      
      if (orderBy) {
        queryParams.push(`$orderby=${orderBy}`);
      }
      
      // Добавляем ограничение на количество возвращаемых элементов для производительности
      queryParams.push('$top=5000');
      
      if (queryParams.length > 0) {
        endpoint += `?${queryParams.join('&')}`;
      }
      
      console.log('Fetching from endpoint:', endpoint);
      
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1
      );
      
      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Error ${response.status}: ${errorText}`);
      }
      
      const result = await response.json();
      return result.value || [];
    } catch (error) {
      console.error('Error in getListItems:', error);
      throw error;
    }
  }
}