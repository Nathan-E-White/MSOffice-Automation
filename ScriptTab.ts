/* Requests */

import("@microsoft/office-js/dist/word-mac-16.00");

//import "@microsoft/office-js/dist/office";
//import "@microsoft/office-js/dist/en-us/office_strings";


const newDoc: Word.DocumentCreated = new Word.DocumentCreated();
const defaultProps: Word.Interfaces.DocumentCreatedData = newDoc.toJSON();
const updateProps: Word.Interfaces.DocumentCreatedUpdateData = {};
const updateOptions: OfficeExtension.UpdateOptions = {};

newDoc.set(updateProps, updateOptions);
newDoc.open();
newDoc.save();

const proxyDocument: OfficeExtension.ClientObject;


const createNewMSWordDocument = (fname: string) => {
    this.set(new Word.DocumentCreated());

}


const reqs = new Array<Request>();
const reqsInfo = new Array<RequestInfo|URL>();
const reqsInit = new Array<RequestInit|null>();


//new Storage()

let reqURL: URL;
let reqInf: RequestInfo;
let reqIni: RequestInit;


const req: Request = new Request(reqInf, reqIni);

const addRequest = (inf: RequestInfo | URL, ini: RequestInit | null): void => {
    reqsInfo.push(inf);
    reqsInit.push(ini);
    reqs.push(new Request(inf, ini));
}