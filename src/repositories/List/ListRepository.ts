import { IPropertyPaneDropdownOption } from "@microsoft/sp-property-pane";
import "@pnp/sp/items";
import "@pnp/sp/lists";
import { IListInfo } from "@pnp/sp/lists";
import "@pnp/sp/lists/web";
import "@pnp/sp/site-groups/web";
import "@pnp/sp/site-users/web";
import "@pnp/sp/sites";
import "@pnp/sp/webs";
import { SharePointRepository } from "../SharePointRepository";

export class ListRepository extends SharePointRepository {
  public getLists = async (
    patternMatch?: string
  ): Promise<IPropertyPaneDropdownOption[]> => {
    try {
      const spListData: IListInfo[] = await this._sp.web.lists
        .select("Title")
        .orderBy("Title")();

      if (!spListData.length) {
        throw new Error("No lists found");
      }

      //If no pattern - return all lists
      if (!patternMatch) {
        return spListData.map(list => {
          return { key: list.Title, text: list.Title };
        });
        //Otherwise return only lists with titles that match the pattern
      } else {
        return spListData
          .filter(list => list.Title.match(patternMatch))
          .map(list => {
            return { key: list.Title, text: list.Title };
          });
      }
    } catch (error) {
      console.error(error);
      return [];
    }
  };
}
