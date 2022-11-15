import { WebPartContext } from "@microsoft/sp-webpart-base";
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/files";
import "@pnp/sp/folders";

export abstract class SharePointRepository {
  protected _sp: any;

  constructor(context: WebPartContext) {
    this._sp = spfi().using(SPFx(context as any));
  }
}
