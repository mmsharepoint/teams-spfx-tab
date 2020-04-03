import * as React from 'react';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import styles from './MeDemo.module.scss';
import GraphController from '../../../controller/GraphController';
import User from './User';
import { IMeDemoProps } from './IMeDemoProps';
import { IMeDemoState } from './IMeDemoState';
import { PersonaPresence } from 'office-ui-fabric-react/lib/Persona';

export default class MeDemo extends React.Component<IMeDemoProps, IMeDemoState> {
  private graphController: GraphController;
  
  constructor(props) {
    super(props);
    this.state = {
      currentUser: null
    };
    this.graphController = new GraphController();
      this.graphController.init(this.props.msGraphClientFactory)
        .then((controllerReady) => {
          if (controllerReady) {
            this.graphClientReady();
          }        
        });
  }
  
  public render(): React.ReactElement<IMeDemoProps> {
    return (
      <div className={ styles.meDemo }>
        <Icon iconName={`${this.props.isMicrosoftTeams?'TeamsLogo':'SharepointLogo'}`} />
        {typeof this.state.currentUser !== 'undefined' && this.state.currentUser !== null &&
        <User currentUser={this.state.currentUser} />}
      </div>
    );
  }

  private graphClientReady = async () => {
    let user = await this.graphController.getMyUser();
    const photoUrl = await this.graphController.getUserPhotoBlobUrl();
    user.photo = photoUrl;
    const presence = await this.graphController.getUserPresence();
    switch (presence) {
      case 'Available':
      case 'AvailableIdle':
        user.presence = PersonaPresence.online;
        break;
      case 'Away':
      case 'BeRightBack':
        user.presence = PersonaPresence.away;
        break;
      case 'Busy':
      case 'BusyIdle':
        user.presence = PersonaPresence.blocked;
        break;
      case 'DoNotDisturb':
        user.presence = PersonaPresence.dnd;
        break;
      case 'Offline':
        user.presence = PersonaPresence.offline;
        break;
      default:
        user.presence = PersonaPresence.none;
        break;
    }        
    this.setState((prevState: IMeDemoState, props:IMeDemoProps) => {
      return {
        currentUser: user
      };
    });
  }
}
