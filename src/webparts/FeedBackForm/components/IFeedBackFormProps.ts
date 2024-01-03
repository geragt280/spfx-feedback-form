
import { DisplayMode } from '@microsoft/sp-core-library';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPFI } from '@pnp/sp';
export interface IFeedBackFormProps {
  title: string;
  displayMode: DisplayMode;
  updateProperty: (value: string) => void;
  sp:SPFI;
  SupportType: string;
  context: WebPartContext;
  FeedbackListID:string;
  colComments: string;
  colSupportType: string;
  colTitle: string;
  headingBackColor: string;
  enableReEnterFormLink: boolean;
}
