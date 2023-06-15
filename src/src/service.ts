import { SPFI } from '@pnp/sp';
import { getSP } from './opspnp';
import { PagedItemCollection } from '@pnp/sp/items';

export class SPInstance{

    private _sp : SPFI;
    constructor(){
        this._sp = getSP();
    }
    public getListItems = async (_listName: string): Promise<any> => {
        let items: any[] = [];
        let itemsCheck: PagedItemCollection<any[]> = undefined;
        do{
        if(!itemsCheck) itemsCheck = await this._sp.web.lists.getByTitle(_listName).items.select('LinkFilename','FileDirRef','FileRef','FileSystemObjectType','Editor/Title','Author/Title').expand("Editor","Author").top(4000).getPaged();
        else itemsCheck = await itemsCheck.getNext();
        if(itemsCheck.results.length > 0){
            items = items.concat(itemsCheck.results)
        }
        } while (itemsCheck.hasNext)
        return await items;
    }
}