declare interface IFeedBackFormWebPartStrings {
  Title: string;
  description: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppLocalEnvironmentOffice: string;
  AppLocalEnvironmentOutlook: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  AppOfficeEnvironment: string;
  AppOutlookEnvironment: string;
  UnknownEnvironment: string;
  lbldrpFeedbackList: string;
  listPickerFieldId: string;
  lblDrpParticipantColumns: string;
  lblDrpCommentsColumn: string;
  lblDrpFeedbackTypeColumn: string;
  enableReEnterFormLink: string;
}

declare module "FeedBackFormWebPartStrings" {
  const strings: IFeedBackFormWebPartStrings;
  export = strings;
}
