import { IPropertyPaneDropdownOption } from '@microsoft/sp-property-pane';
import { IGroup } from '../models';

/**
 * Utils
 * @class
 */
export class Utils {

    /** Convert array of O365 groups into an array of DropDown options
     * @param grp Array of O365 groups
     * @return DropDown options of O365 groups
     */
    public static convertGrpToOptions(grp: Array<IGroup>): Array<IPropertyPaneDropdownOption> {
        var options: Array<IPropertyPaneDropdownOption> = new Array<IPropertyPaneDropdownOption>();
        if (grp && grp.length > 0) {
            grp.map((g: IGroup, idx: number) => {
                options.push({ key: idx, text: g.displayName });
            });
        }
        return options;
    }
}