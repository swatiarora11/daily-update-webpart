import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration } from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import { DailyUpdateWebPartPropertyPane } from "./DailyUpdateWebPartPropertyPane";
import { IPropertyFieldSite } from "@pnp/spfx-property-controls/lib/PropertyFieldSitePicker";

import { Services } from "../../services";
import { IGroup } from "../../models";

import styles from './DailyUpdateWebPart.module.scss';
import { DailyUpdate } from './DailyUpdate';

export interface IDailyUpdateWebPartProps {
  title: string;
  description: string;
  iconProperty: string;
  context: WebPartContext;
  adminUpn: string;
  selectedGroupIdx: number;
  teamsAppId: string;
  selectedSite: IPropertyFieldSite[];
  selectedList: any;
}

let groups: IGroup[] = undefined;
let _this: any;

export default class DailyUpdateWebPart extends BaseClientSideWebPart<IDailyUpdateWebPartProps> {
  private _deferredPropertyPane: DailyUpdateWebPartPropertyPane | undefined;
  private updates: DailyUpdate = undefined;
  private services: Services = undefined;

  private forcedRefreshCount = 3;
  private timeoutId: number = 0;
  private setNextUpdateTimeOut(today: Date) {
    if(this.properties.selectedList) {
      clearTimeout(this.timeoutId);
      if(this.forcedRefreshCount <= 0) {
        this.timeoutId = setTimeout(this.render, this.updates.getNextDayTimeAt(6, 30).getTime() - today.getTime()); 
        this.forcedRefreshCount = 3; 
      }
      else {
        this.timeoutId = setTimeout(this.render, 5000); 
        this.forcedRefreshCount--;
      } 
    }
  }

  public async onInit(): Promise<void> {
    _this = this;

    this.services = new Services(this.context);
    await this.services.init();

    this.updates = new DailyUpdate(this.services);
    groups = await this.services.getGroups();

    return Promise.resolve();
  }

  public OnRefreshButtonClick() {
    this.render();
  }

  public OnFeedButtonClick() {
    this.updates.sendTeamsFeed(groups[this.properties.selectedGroupIdx], this.properties.teamsAppId, this.properties.adminUpn);
  }

  public render(): void {
    console.log("Rendering daily web part " + this.forcedRefreshCount);
    var today = new Date();
    var site = this.properties.selectedSite;

    if(site) {
      this.updates._getDailyUpdate(today, site[0]?.id ?? "", this.properties.selectedList?.id ?? "");
    }

    var feedButton = "";
    if(this.properties.adminUpn) {
      feedButton = `<button id="btnSendFeed" type="submit" class="${ styles.button }">Send Feed</button>`;
    }

    this.domElement.innerHTML = `
        <div class="${ styles.dailyUpdate }">
          <table noborder>
            <td width='25%'><img align = "right" src="${ this.properties.iconProperty}" alt="Chairman Image" style="width:200px;height:auto;"></td>
            <td>
              <div class="${ styles.container }">
                <div class="${ styles.row }">
                  <div class="${ styles.column }">
                    <span class="${ styles.title }">${ "[" + this.updates.formattedDate + "] " }</span> 
                    <span class="${ styles.title }">${ this.properties?.title}</span>
                    <p class="${ styles.subTitle }">${ "~" + this.updates.author }</p>
                    <p class="${ styles.title }">${escape('\"' + this.updates.description + '\"')}</p>
                    ${ feedButton }
                    </div>           
                </div>
              </div>
            </td>
          </table>
        </div>`;

      if(this.properties.adminUpn) {
        document.getElementById('btnSendFeed').addEventListener('click',()=>this.OnFeedButtonClick()); 
      }

      this.setNextUpdateTimeOut(today);
  }

  public get title(): string {
    return this.properties.title;
  }

  protected get iconProperty(): string {
    return this.properties.iconProperty || require("../../assets/SharePointLogo.svg");
  }

  protected get adminUpn(): string {
    return this.properties.adminUpn;
  }

  protected get selectedGroupIdx(): number {
    return this.properties.selectedGroupIdx;
  }

  protected get teamsAppId(): string {
    return this.properties.teamsAppId;
  }

  protected get selectedSite(): IPropertyFieldSite[] {
    return this.properties.selectedSite;
  }

  protected get selectedList(): any {
    return this.properties.selectedList;
  }

  protected onPropertyPaneFieldChanged = async (propertyPath: string, oldValue: any, newValue: any) => {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);

    if (propertyPath == "selectedSite" && newValue !== oldValue) {
      const site: IPropertyFieldSite[] = newValue as IPropertyFieldSite[];
      this.properties.selectedList = undefined;
      this.properties.selectedSite = site;
    }
    if (propertyPath == "selectedList" && newValue != oldValue) {
      this.properties.selectedList = newValue;
    }
    if (propertyPath == "adminUpn" && newValue != oldValue) {
      this.properties.adminUpn = newValue;
    }
    if (propertyPath == "selectedGroupIdx" && newValue != oldValue) {
      this.properties.selectedGroupIdx = newValue;
    }
    if (propertyPath == "teamsAppId" && newValue != oldValue) {
      this.properties.teamsAppId = newValue;
    }

    this.context.propertyPane.refresh();
    this.render();
  }

  protected async loadPropertyPaneResources(): Promise<void> {
    const component = await import(

      "./DailyUpdateWebPartPropertyPane"
    );

    this._deferredPropertyPane = new component.DailyUpdateWebPartPropertyPane(
      this.context,
      this.properties,
      groups,
      this.onPropertyPaneFieldChanged
    );
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
  
  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return this._deferredPropertyPane!.getPropertyPaneConfiguration();
  }
}
