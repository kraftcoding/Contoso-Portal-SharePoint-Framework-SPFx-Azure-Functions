import { SPFI, spfi } from '@pnp/sp';
import { SPFx } from '@pnp/sp/behaviors/spfx';

import { WebPartContext } from '@microsoft/sp-webpart-base';

import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';


export interface IPeriodItem {
  Id: number;
  title?: string;
  codigo?: string;
  descripcion?: string;
  StartDate?: Date;
  EndDate?: Date;
  activo?: boolean;
}

type RawItem = {
  Id: number;
  Title?: string;
  codigo?: string;
  descripcion?: string;
  StartDate?: string; // ISO string desde SP
  EndDate?: string;    // ISO string desde SP
  activo?: boolean;
};

export class PeriodRepository {
  private sp: SPFI;
  private listTitle = 'Period';


  constructor(private context: WebPartContext) {
    if (!context) throw new Error('SPFx context es undefined');
    this.sp = spfi().using(SPFx(this.context));
  }


  private toModel = (r: RawItem): IPeriodItem => ({
    Id: r.Id,
    title: r.Title,
    codigo: r.codigo,
    descripcion: r.descripcion,
    StartDate: r.StartDate ? new Date(r.StartDate) : undefined,
    EndDate: r.EndDate ? new Date(r.EndDate) : undefined,
    activo: r.activo
  });

  // asegurar longitud (255) y evitar nulos
  private toTitleFromDescripcion(data: Partial<IPeriodItem>): string | undefined {
    const t = (data.descripcion ?? '').trim();
    return t ? t.substring(0, 255) : undefined; // SharePoint Title máx. 255 chars
  }

  public async getAll(): promise<IPeriodItem[]> {
    const items = await this.sp.web.lists.getByTitle(this.listTitle).items
      .orderBy("StartDate")
      .select('Id', 'codigo', 'descripcion', 'StartDate', 'EndDate', 'activo')<RawItem[]>();
    return items.map(this.toModel);
  }

  public async add(data: Partial<IPeriodItem>): promise<number> {
    const res = await this.sp.web.lists.getByTitle(this.listTitle).items.add({
      Title: this.toTitleFromDescripcion(data),
      codigo: data.codigo,
      descripcion: data.descripcion,
      StartDate: data.StartDate ? data.StartDate.toISOString() : undefined,
      EndDate: data.EndDate ? data.EndDate.toISOString() : undefined,
      activo: data.activo
    });
    return res.data?.Id;
  }

  public async update(id: number, data: Partial<IPeriodItem>): promise<void> {
    await this.sp.web.lists.getByTitle(this.listTitle).items.getById(id).update({
      Title: this.toTitleFromDescripcion(data),
      codigo: data.codigo,
      descripcion: data.descripcion,
      StartDate: data.StartDate ? data.StartDate.toISOString() : undefined,
      EndDate: data.EndDate ? data.EndDate.toISOString() : undefined,
      activo: data.activo
    });
  }

  public async remove(id: number): promise<void> {
    await this.sp.web.lists.getByTitle(this.listTitle).items.getById(id).delete();
  }

  public async checkLookupFieldsInUse(id: number, field:string, listUrl: string):promise<boolean>{
    const list = this.sp.web.getList(listUrl);
    const camlQuery = {
      ViewXml: 
        `<View Scope="RecursiveAll">
          <Query>
            <Where>
              <And>
                <Geq>
                  <FieldRef Name='ID' />
                  <Value Type='Counter'>1</Value>
                </Geq>
                <Eq>
                  <FieldRef Name='${field}' LookupId='TRUE' />
                  <Value Type='Lookup'>${id}</Value>
                </Eq>
              </And>
            </Where>
            <OrderBy>
              <FieldRef Name='ID' Ascending='TRUE' />
            </OrderBy>
          </Query>
          <ViewFields>
              <FieldRef Name='ID' />
            </ViewFields>
          <RowLimit Paged='TRUE'>1</RowLimit>
        </View>`
    }

    try{
      const response = await list.renderListDataAsStream(camlQuery);
      if(response && response.Row && response.Row.length > 0){
        return true;
      }else{
        return false;
      }
    }catch(ex){
        console.error(ex);
        console.error("error obteniedo datos de la lista")
        return true;
    }
  }
}