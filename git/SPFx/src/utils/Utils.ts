import { find } from "@microsoft/sp-lodash-subset";
import { Term } from "../models/ITag";
import { PnPClientStorage } from '@pnp/core';
import { EnvConfig } from "./EnvConfig";
import { OrderTypes } from "../service/BackendServiceModels/EventModels";

export function getTermLabel(terms: Term[], id: string): string {
  const term: Term | undefined = find(terms, (term: Term) => term.key === id);
  return term ? term.text : '';
}

export function getTerm(terms: Term[], id: string): Term | undefined {
  const term: Term | undefined = find(terms, (term: Term) => term.key === id);
  return term;
}

export function IsNullOrECNTy(s:string| null | undefined | Record<string, any>){
  return s === null || s === undefined || s === '' ||!(Object.keys(s)?.length > 0);
}



export function groupBy<T> (array: Array<T>, property: (x: T) => string): { [key: string]: Array<T> } {
  return array.reduce((memo: { [key: string]: Array<T> }, x: T) => {
    if (!memo[property(x)]) {
      memo[property(x)] = [];
    }
    memo[property(x)].push(x);
    return memo;
  }, {});
}

export function getOrderFormatted(order: number, type: OrderTypes): string {
  order = order ?? 1;
  switch(type){
    case OrderTypes.Alphabetical:
      return String.fromCharCode(64 + order);
    case OrderTypes.Roman:
      return numberToRoman(order);
    default:
    case OrderTypes.Numeric:
      return order.toString();
  }
}

export function numberToRoman(num: number): string{
  const mapa: {[x: string] : number } = { L: 50, XL: 40, X: 10, IX: 9, V: 5, IV: 4, I: 1 };
  let resultado = "";

  for (const simbolo in mapa) {  // eslint-disable-line
      while (num >= mapa[simbolo]) {
          resultado += simbolo;
          num -= mapa[simbolo];
      }
  }

  return resultado;
}
 

export function getTimeFormatted(date?: Date): string {
  return date?.toLocaleTimeString(undefined, { hour: '2-digit', minute: '2-digit' }) || '';
}

export function getDateFormatted(date: Date | undefined): string {
  return date ? date.toLocaleDateString(undefined, { day: '2-digit', month: '2-digit', year: 'numeric' }) : '';
}

const _cacheBaseKey = 'spfx_';
const _storage = new PnPClientStorage();
export async function cacheGetOrPut<T>(key: string, getter: () => promise<T>, minutesToExpire?: number): promise<T> {
  let expiration = new Date();
  expiration.setMinutes(expiration.getMinutes() + (minutesToExpire ?? parseInt(EnvConfig.FrontendCacheMinutes)));
  return _storage.session.getOrPut<T>(
    `${_cacheBaseKey}${key}`,
    async () => {
      return getter();
    },
    expiration,
  );
}

export function cacheRemove(key: string): void {
  _storage.session.delete(`${_cacheBaseKey}${key}`);
}

export function sortByprodp<T>(array: any[], prodp: string): T[] {
  return array.sort((a, b) => {
    const nameA = a[prodp].toUpperCase();
    const nameB = b[prodp].toUpperCase();
    if (nameA < nameB) return -1;
    if (nameA > nameB) return 1;
    return 0;
  });
}

export function getTime(date?: string): number {
  return date != null ? new Date(date).getTime() : 0;
}