import * as React from 'react';
import {
  IPersonaSharedProps,
  Persona,
  PersonaSize,
  PersonaPresence
} from 'office-ui-fabric-react/lib/Persona';
import styles from './User.module.scss';
import { IUserProps } from './IUserProps';

export default class User extends React.Component<IUserProps, {}> {  
  constructor(props) {
    super(props);    
  }
  
  public render(): React.ReactElement<IUserProps> {
    const userProfile: IPersonaSharedProps = {
      text: this.props.currentUser.displayName,
      secondaryText: this.props.currentUser.jobTitle,
      tertiaryText: this.props.currentUser.mobilePhone      
    };
    return (
      <div className={styles.user}>
        <div className={styles.row}>
          <div className={styles.fullCol}>
            <Persona {...userProfile}
                      size={PersonaSize.size100}
                      presence={this.props.currentUser.presence}
                      imageAlt={this.props.currentUser.displayName}
                      imageUrl={this.props.currentUser.photo}
                      coinSize={110}
                    />
          </div>
        </div>               
        <div className={styles.row}>
          <div className={styles.leftCol}><label>Firstname</label></div>
          <div className={styles.rightCol}>{this.props.currentUser.givenName}</div>
        </div>
        <div className={styles.row}>
          <div className={styles.leftCol}><label>Lastname</label></div>
          <div className={styles.rightCol}>{this.props.currentUser.surname}</div>
        </div>
        <div className={styles.row}>
          <div className={styles.leftCol}><label>Mail</label></div>
          <div className={styles.rightCol}>{this.props.currentUser.mail}</div>
        </div>        
        <div className={styles.row}>
          <div className={styles.leftCol}><label>Office Location</label></div>
          <div className={styles.rightCol}>{this.props.currentUser.officeLocation}</div>
        </div>   
        <div className={styles.row}>
          <div className={styles.leftCol}><label>ID</label></div>
          <div className={styles.rightCol}>{this.props.currentUser.id}</div>
        </div>              
      </div>
    );
  }
}