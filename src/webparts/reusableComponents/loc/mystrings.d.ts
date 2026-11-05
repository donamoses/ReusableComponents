declare interface IReusableComponentsWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  WebpartTitle: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppLocalEnvironmentOffice: string;
  AppLocalEnvironmentOutlook: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  AppOfficeEnvironment: string;
  AppOutlookEnvironment: string;
  UnknownEnvironment: string;
  DivStyle: string;
  DivStyleOn: string;
  DivStyleOff: string;
}

declare module 'ReusableComponentsWebPartStrings' {
  const strings: IReusableComponentsWebPartStrings;
  export = strings;
}
