import { ISPService } from "../../../service/Service";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IDirectLinksprops {
  spService: ISPService;
  context: WebPartContext;
  title: string;
  slidesToShow: number;
  slidesToScroll: number;
}