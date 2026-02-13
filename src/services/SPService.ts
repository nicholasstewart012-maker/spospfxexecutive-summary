import { WebPartContext } from "@microsoft/sp-webpart-base";
import { spfi, SPFx, SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { IEvent, EventFields } from "../common/IEvent";
import { Web } from "@pnp/sp/webs";

export class SPService {
    private _sp: SPFI;

    constructor(context: WebPartContext) {
        this._sp = spfi().using(SPFx(context));
    }

    public async getEvents(siteUrl: string, listName: string): Promise<IEvent[]> {
        try {
            if (!listName || !siteUrl) {
                return [];
            }

            const web = Web([this._sp.web, siteUrl]);
            const items: any[] = await web.lists
                .getByTitle(listName)
                .items
                .select(...EventFields)
                .top(500)
                .orderBy("EventDate", true)();

            return items.map(item => ({
                Id: item.Id,
                Title: item.Title,
                EventDate: item.EventDate,
                EndDate: item.EndDate,
                Description: item.Description,
                Category: item.Category,
                CategoryColor: item.CategoryColor,
                TargetAudience: item.TargetAudience,
                Location: item.Location,
                Department: item.Department,
                Contact: item.Contact,
                SubjectMatterExpert: item.SubjectMatterExpert,
                Prework: item.Prework,
                RegistrationDate: item.RegistrationDate,
                StartTimeZoneCST: item.StartTimeZoneCST,
                EndTimeZoneCST: item.EndTimeZoneCST,
                StartTimeZoneEST: item.StartTimeZoneEST,
                EndTimeZoneEST: item.EndTimeZoneEST
            }));
        } catch (error: any) {
            console.error("Error fetching events:", error);
            throw error;
        }
    }
}
