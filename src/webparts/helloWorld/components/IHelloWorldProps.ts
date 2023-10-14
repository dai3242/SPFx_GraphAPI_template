import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IHelloWorldProps {
  description?: string;
  test?:string;
  siteName?:string;
}

export interface IHelloWorldWebPartProps {
  context: WebPartContext;
  properties: IHelloWorldProps;
}