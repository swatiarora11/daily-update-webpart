import { ListItem, TeamsAppDefinition } from "@microsoft/microsoft-graph-types";
import { BaseComponentContext } from "@microsoft/sp-component-base";
import { MSGraphClient } from "@microsoft/sp-http";
import { IGroup, IGroupCollection, IUser, ITeamsAppCollection } from "../models";

export class Services {
  private _context: BaseComponentContext = undefined;
  private _msGraphClient = undefined;

  constructor(context: BaseComponentContext) {
    this._context = context;
  }

  public init = async  () => {
    this._msGraphClient = await this._context.msGraphClientFactory.getClient();
  }

  public getUser(upn: string): Promise<any> {
    return new Promise<any>((resolve, reject) => {
      try {
        this._context.msGraphClientFactory
          .getClient()
          .then((client: MSGraphClient) => {
            client
              .api(`/users/${upn}`)
              .get((error: any, user: IUser, rawResponse: any) => {
                resolve(user);
              });
          });
      } catch (error) {
        console.error(error);
      }
    });
  }

  public getGroups(): Promise<IGroup[]> {
    return new Promise<IGroup[]>((resolve, reject) => {
      try {
        // Prepare the output array
        var o365groups: Array<IGroup> = new Array<IGroup>();

        this._context.msGraphClientFactory
          .getClient()
          .then((client: MSGraphClient) => {
            client
              .api("/groups?$filter=groupTypes/any(c:c eq 'Unified')")
              .get((error: any, groups: IGroupCollection, rawResponse: any) => {
                // Map the response to the output array
                groups.value.map((item: any) => {
                  o365groups.push({
                    id: item.id,
                    displayName: item.displayName,
                    description: item.description,
                    visibility: item.visibility
                  });
                });

                resolve(o365groups);
              });
          });
      } catch (error) {
        console.error(error);
      }
    });
  }

  public getGroup(groupId: string): Promise<any> {
    return new Promise<any>((resolve, reject) => {
      try {
        this._context.msGraphClientFactory
          .getClient()
          .then((client: MSGraphClient) => {
            client
              .api(`/groups/${groupId}`)
              .get((error: any, group: any, rawResponse: any) => {
                resolve(group);
              });
          });
      } catch (error) {
        console.error(error);
      }
    });
  }

  public getGroupLink(groupId: string): Promise<any> {
    return new Promise<any>((resolve, reject) => {
      try {
        this._context.msGraphClientFactory
          .getClient()
          .then((client: MSGraphClient) => {
            client
              .api(`/groups/${groupId}/sites/root/weburl`)
              .get((error: any, group: any, rawResponse: any) => {
                resolve(group);
              });
          });
      } catch (error) {
        console.error(error);
      }
    });
  }

  public getGroupThumbnail(groupId: string): Promise<any> {
    return new Promise<any>((resolve, reject) => {
      try {
        this._context.msGraphClientFactory
          .getClient()
          .then((client: MSGraphClient) => {
            client
              .api(`/groups/${groupId}/photos/48x48/$value`)
              .responseType('blob')
              .get((error: any, group: any, rawResponse: any) => {
                resolve(window.URL.createObjectURL(group));
              });
          });
      } catch (error) {
        console.error(error);
      }
    });
  }

  public getGroupMembers(groupId: string): Promise<string[]> {
    return new Promise<any>((resolve, reject) => {
      try {
        this._context.msGraphClientFactory
        .getClient()
        .then((client: MSGraphClient) => {
          client
            .api(`/groups/${groupId}`)
            .expand('members')
            .get((error: any, groupmembers: any, rawResponse: any) => {
                let membersAndOwners: string[] = [];
                if (groupmembers.members) {
                  groupmembers.members.forEach(member => {
                    membersAndOwners.push(member.userPrincipalName);
                  });
                }
                resolve(membersAndOwners);
            });
          });
      } catch (error) {
        console.error(error);
      }
    });
  }

  public getListItem = (
    siteId: string,
    listId: string,
    itemId: string
  ): Promise<any> => {
    return new Promise<any>((resolve, reject) => {
      try {
        this._context.msGraphClientFactory
          .getClient()
          .then((client: MSGraphClient) => {
            client
              .api(`/sites/${siteId}/lists/${listId}/items/${itemId}`)
              .get((error: any, item: ListItem, rawResponse: any) => {
                resolve(item);
              });
          });
      } catch (error) {
        console.error(error);
      }
    });
  }

  public getListItemByQuery = (
    siteId: string,
    listId: string,
    query: string
  ): Promise<any> => {
    return new Promise<any>((resolve, reject) => {
      try {
        this._context.msGraphClientFactory
          .getClient()
          .then((client: MSGraphClient) => {
            client
              .api(`/sites/${siteId}/lists/${listId}/items${query}`)
              .get((error: any, results: any, rawResponse: any) => {
                resolve(results?.value[0]);
              });
          });
      } catch (error) {
        console.error(error);
      }
    });
  }

  public getTeamsAppInstance = (
    userId: string,
    teamsAppId: string
  ): Promise<any> => {
    return new Promise<any>((resolve, reject) => {
      try {
        this._context.msGraphClientFactory
          .getClient()
          .then((client: MSGraphClient) => {
            client
              .api(`https://graph.microsoft.com/v1.0/users/${userId}/teamwork/installedApps?$expand=teamsApp&$filter=(teamsApp/id eq '${teamsAppId}')`)
              .get((error: any, instances: ITeamsAppCollection, rawResponse: any) => {
                resolve(instances.value[0]);
              });
          });
      } catch (error) {
        console.error(error);
      }
    });
  }

  public sendActivityFeedUser = (
    feedUserId: string,
    appInstallationId: string,
    feedPreviewText: string
  ): Promise<any> => {
    return new Promise<any>((resolve, reject) => {
      try {
        const endpoint: string = `https://graph.microsoft.com/v1.0/users/${feedUserId}/teamwork/installedApps/${appInstallationId}`;
        const notificationBody: any = {
          topic: {
            source: "entityUrl",
            value: endpoint,
          },
          activityType: "readThisRequired",
          previewText: {
            content: feedPreviewText,
          }
        };

        this._context.msGraphClientFactory
          .getClient()
          .then((client: MSGraphClient) => {
            client
              .api(`/users/${feedUserId}/teamwork/sendActivityNotification`)
              .post(notificationBody, (error: any, result: any, rawResponse: any) => {
                resolve(result);
              });
          });
      } catch (error) {
        console.error(error);
      }
    });
  }
}


