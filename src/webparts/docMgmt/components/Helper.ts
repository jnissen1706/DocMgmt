
import { MSGraphClientV3 } from '@microsoft/sp-http';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export default class Helper {
    //Date Format - [2020-10-31T07:00:00Z] to [MM/DD/YYYY]
    /*public static ConvertDateFormate(dateValue: any) {
        let check = (n: any) => { return (Number(n) < 10) ? "0" + n : String(n); }; //Used for leading 0
        let date = new Date(dateValue);
        let delimiter = "/";
        let newDate = check(date.getMonth() + 1) + delimiter + check(date.getDate()) + delimiter + date.getFullYear();
        return newDate;
    }*/

    //Get Query String value
    public static GetQueryStringValue = async(queryStringParameter: string): Promise<string> => {
        try {
            const currentUrl = new URL(window.location.href.toLowerCase());
            if (currentUrl.searchParams.get(queryStringParameter) !== null) {
                return currentUrl.searchParams.get(queryStringParameter);
            }
            return '';
        }
        catch(ex) {
            return '';
        }
    }


    //Get EMail Message details using MS Graph
    public static GetMessageDetails = async(context: WebPartContext): Promise<string> => {
        try {
            const promise: Promise<string> = new Promise<string>(
                (resolve, reject) => {
                    context.msGraphClientFactory
                        .getClient('3')
                        .then((client: MSGraphClientV3): void => {
                            client
                            .api('/me')
                            .get((error: any, response: MicrosoftGraph.User, rawResponse?: any) => {
                                // handle the response
                                console.log(response);
                            })
                        })
                        .catch((errorL: any) => {
                            console.log(``);
                        });
                }
            );
    
            return promise;

        }
        catch(ex) {
            console.log(`Error`)
        }
    }
}