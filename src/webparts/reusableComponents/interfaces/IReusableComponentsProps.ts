import { WebPartContext } from "@microsoft/sp-webpart-base";
import type { IReusableComponentsWebPartProps } from "../ReusableComponentsWebPart";

export interface IReusableComponentsProps {
  context: WebPartContext;
  wpproperties: IReusableComponentsWebPartProps;
}
