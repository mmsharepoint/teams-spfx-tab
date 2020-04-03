import { MSGraphClient, MSGraphClientFactory } from '@microsoft/sp-http';
import { GraphError } from '@microsoft/microsoft-graph-client/lib/src/common';
import { IUser } from '../model/IUser';
import { PersonaPresence } from 'office-ui-fabric-react/lib/Persona';

export default class GraphController {
    private client: MSGraphClient;

    public init(graphFactory: MSGraphClientFactory): Promise<boolean> {
        return graphFactory
            .getClient()
            .then((client: MSGraphClient) => {
            this.client = client;
                return true;
            })
            .catch((error) => {
                return false;
            });  
    }

    public getMyUser(): Promise<IUser> {
        return this.getMyUserData();
    }

    public getUserPhotoBlobUrl(): Promise<string> {
        return this.getMyUserPhoto();
    }

    public getUserPresence(): Promise<string> {
        return this.getMyUserPresence();
    }

    private getMyUserData(): Promise<IUser> {        
        return this.client
            .api('me')
            .version('v1.0')
            .get()
            .then((response) => {
                const user: IUser = {
                        displayName: response.displayName,
                        givenName: response.givenName,
                        surname: response.surname,
                        id: response.id,
                        jobTitle: response.jobTitle,
                        mail: response.mail,
                        mobilePhone: response.mobilePhone,
                        officeLocation: response.officeLocation,
                        preferredLanguage: response.preferredLanguage,
                        userPrincipalName: response.userPrincipalName,
                        photo: null,
                        presence: PersonaPresence.none
                };
                return user;
            });
    }

    private getMyUserPhoto(): Promise<string> {        
        return this.client
            .api('/me/photos/120x120/$value') // 48x48, 64x64, 96x96, 120x120, 240x240, 360x360, 432x432, 504x504, and 648x648
            .version('v1.0')                        
            .responseType('blob')
            .get()
            .then((photoResponse: any) => {
                const blobUrl = window.URL.createObjectURL(photoResponse);                   
                return blobUrl;
            });
    }

    private getMyUserPresence(): Promise<string> {        
        return this.client
            .api('/me/presence') // Available, AvailableIdle, Away, BeRightBack, Busy, BusyIdle, DoNotDisturb, Offline, PresenceUnknown
            .version('beta')                        
            .get()
            .then((response: any) => {                                  
                return response.availability;
            });
    }
}