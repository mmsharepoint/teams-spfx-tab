import { PersonaPresence } from "office-ui-fabric-react/lib/Persona";

export interface IUser {
    displayName: string;
    givenName: string;
    jobTitle: string;
    mail: string;
    mobilePhone: string;
    officeLocation: string;
    preferredLanguage: string;
    surname: string;
    userPrincipalName: string;
    id: string;
    photo: string;
    presence: PersonaPresence;
}