import { IGroup } from "../../models";
import { Services } from "../../services";

export class DailyUpdate {
  private services: Services = undefined;

  public description: string;
  public author: string;
  public formattedDate: string;

  constructor(services: Services) {
    this.services = services;
  }
   
  public getNextDayTimeAt(hour: number, mins: number) {
    var t = new Date();
    t.setDate(t.getDate()+1);
    t.setHours(hour);
    t.setMinutes(mins);
    t.setSeconds(0);
    t.setMilliseconds(0);
    return t;
  }

  public async _getDailyUpdate(today: Date, siteId: string, listId: string) {
    const months = ["January", "February", "March","April", "May", "June", "July", "August", "September", "October", "November", "December"];
    this.formattedDate = today.getDate() + " " + months[today.getMonth()] + " " + today.getFullYear();

    if(siteId != "" && listId != "") {
      var Startdate = today.toISOString().substring(0, 10) + "T00:00:00.000Z";
      var Enddate =  today.toISOString().substring(0, 10) + "T23:00:00.000Z";
      var filterQuery = "?top=1&filter=fields/Live ge '" + Startdate + "' and fields/Live le '" + Enddate + "'";

      try {
        const currItem = await this.services.getListItemByQuery(siteId, listId, filterQuery);
        if(currItem != undefined) {
          const listItem = await this.services.getListItem(siteId, listId, currItem.id);
          const { Title, Update } = listItem.fields;
          this.description = Update;
          this.author = Title;
        }
      } catch (error) {
        console.log(error);
      }
    }
  }

  public async sendTeamsFeed(group: IGroup, teamsAppId: string, adminUpn: string) {
    if(group && teamsAppId != "" && adminUpn != "") {      
      const upns = await this.services.getGroupMembers(group?.id);
      for(var index: number=0; upns && index < upns.length; index++) {
        if(upns[index].toLowerCase() !== adminUpn.toLowerCase()) {
          const teamsAppUser = await this.services.getUser(upns[index]);
          if(teamsAppUser) {
            const teamsAppInstance = await this.services.getTeamsAppInstance(teamsAppUser?.id, teamsAppId); 
            if(teamsAppInstance) { 
              console.log("sending feed to " + teamsAppUser?.displayName);
              this.services.sendActivityFeedUser(teamsAppUser?.id, teamsAppInstance?.id, "click here to see the latest");
            }
          }
        }
      } 
    }
  }
}