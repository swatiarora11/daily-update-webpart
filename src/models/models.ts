export interface IUser {
    id: string;
    displayName: string;
    userPrincipalName?: string;
}

export interface IUserCollection {
    value: IUser[];
}

export interface ITeamsApp {
    id: string;
    displayName: string;
}

export interface ITeamsAppInstallation {
    id: string;
    teamsApp?: ITeamsApp;
}

export interface ITeamsAppCollection {
    value: ITeamsAppInstallation[];
}

export interface IGroup {
    id: string;
    displayName: string;
    description?: string;
    visibility?: string;
    url?: string;
    thumbnail?: string;
    userRole?: string;
}

export interface IGroupCollection {
    value: IGroup[];
}