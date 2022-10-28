import * as React from 'react';
import styles from './DocMgmt.module.scss';
import { IDocMgmtProps } from './IDocMgmtProps';
import { IDocMgmtState, ComponentState } from './IDocMgmtState';
import { escape } from '@microsoft/sp-lodash-subset';
import Helper from './Helper';
import { Panel, PanelType, Spinner } from 'office-ui-fabric-react';
import { getSP } from "../pnpjsConfig";
import { SPFI } from "@pnp/sp";

export default class DocMgmt extends React.Component<IDocMgmtProps, IDocMgmtState> {
  private _sp: SPFI;

  constructor(props: IDocMgmtProps) { 
    super(props);

    this.state = {
      errorMessage: '',
      disaplyState: ComponentState.loadingSpinner
    }
    
    this._sp = getSP();
  }

  public async componentDidMount(): Promise<void> {
    try {
      //Check Query String for Message ID (MID) - This is to directly display the panel
      const messageID = await Helper.GetQueryStringValue("mid");
      if(messageID !== "") {
        //Get Email Message details based on MID
        const valu =  Helper.GetMessageDetails(this.props.context);
        if(valu !== null) {
          this.setState({
            disaplyState: ComponentState.archiveAttachments
          });
        }
        console.log(valu);
      }
      else {
        //Query String is empty. Display static message
        this.setState({
          disaplyState: ComponentState.noMessageID
        });
      }
    }
    catch(ex) {
      console.error(`Error in ComponentDidMount - ${ex}`);
    }
  }

  // error handlar method 
  private _onError(message: string): void {
    this.setState({
      disaplyState: ComponentState.error,
      errorMessage: message
    });
  }

  public render(): React.ReactElement<IDocMgmtProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage
    } = this.props;

    return (
      <section className={styles.docMgmt}>
        {{
          0: /*Loading Message*/
            <div>
              <Panel
                isOpen={true}
                type={PanelType.smallFluid}
                hasCloseButton={false}>
                <div>
                  <Spinner />
                </div>
              </Panel>
            </div>,
          1: /*Error*/
            <div>
              <Panel
                isOpen={true}
                type={PanelType.smallFluid}
                hasCloseButton={false}>
                <div>
                <span>{this.state.errorMessage}</span>
                </div>
              </Panel>
            </div>,
          2: /*Archive Email Message/Attachments*/
            <div>
              <Panel
                isOpen={true}
                type={PanelType.smallFluid}
                hasCloseButton={false}>
                <div>
                  <span>Content here</span>
                </div>
              </Panel>
            </div>,
          3: /*No Message ID*/
            <div>
              <Panel
                isOpen={true}
                type={PanelType.smallFluid}
                hasCloseButton={false}>
                <div>
                <span>Empty Message ID</span>
                </div>
              </Panel>
            </div>
        }[this.state.disaplyState]}
      
        <div className={styles.welcome}>
          <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
          <div>{environmentMessage}</div>
          <div>Web part property value: <strong>{escape(description)}</strong></div>
        </div>
        <div>
          <h4>Learn more about SPFx development:</h4>
          <ul className={styles.links}>
            <li><a href="https://aka.ms/spfx" target="_blank" rel="noreferrer">SharePoint Framework Overview</a></li>
          </ul>
        </div>
      </section>
    );
  }
}
