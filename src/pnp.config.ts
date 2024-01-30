import { spfi, SPFI, SPFx as spSPFx } from "@pnp/sp";
import { LogLevel, PnPLogging } from "@pnp/logging";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";


// eslint-disable-next-line no-var
var _sp: SPFI;

export const getSP = (context?: WebPartContext): SPFI => {
	if (context != null) {
		//You must add the @pnp/logging package to include the PnPLogging behavior it is no longer a peer dependency
		// The LogLevel set's at what level a message will be written to the console
		_sp = spfi().using(spSPFx(context)).using(PnPLogging(LogLevel.Warning));
	}
	return _sp;
};