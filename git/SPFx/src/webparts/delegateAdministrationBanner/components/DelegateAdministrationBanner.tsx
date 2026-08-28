import * as React from 'react';
import styles from './DelegateAdministrationBanner.module.scss';
import type { IDelegateAdministrationBannerprops, IDelegateAdministrationBannerState } from './IDelegateAdministrationBanner';
import { BodyRole, userIsAnyOf } from '../../../service/RoleService';
import { Label, Link } from '@fluentui/react-components';
import strings from 'DelegateAdministrationBannerWebPartStrings';

export default class DelegateAdministrationBanner extends React.Component<IDelegateAdministrationBannerprops, IDelegateAdministrationBannerState> {
  constructor(props: IDelegateAdministrationBannerprops) {
    super(props);
    this.state = {
      showBanner: false
    };
  }

  async componentDidMount(): promise<void> {
    const {context, spService} = this.props;
    const isEditor:boolean= await userIsAnyOf(context, spService, [BodyRole.Scheduler, BodyRole.SchedulerAssistant]);
    this.setState({showBanner: isEditor});
    const hasParentRender = context.domElement.parentElement?.parentElement?.parentElement?.classList?.contains("ControlZone");
    if(hasParentRender){
      const parentRender = context.domElement.parentElement?.parentElement?.parentElement;
      if(parentRender && !isEditor){
        parentRender.style.margin = "0px";
        parentRender.style.padding = "0px"
      }else if(parentRender && isEditor){
        parentRender.style.margin = "0px";
      }
    }
  }

  public render(): React.ReactElement<IDelegateAdministrationBannerprops> {
    const {
      showBanner
    } = this.state;
    return (
      <div className={showBanner ? styles.delegateAdministrationBanner : styles.noShow}>
        {showBanner &&
          <div>
            {this.props.componentTitle &&
              <div className={styles.wpTitle}>{this.props.componentTitle}</div>
            }
            <div className={styles.content}>
              <Label className={styles.msg}>{strings.Msg}</Label>
              <Link 
              className={styles.link} 
              //href={this.props.context.pageContext.web.serverRelativeUrl+'/sitepages/Administracion.aspx'} 
              target='_blank'
              onClick={(e) => {
                e.stopprodpagation(); 
                e.preventDefault()
                window.open(this.props.context.pageContext.web.serverRelativeUrl+'/sitepages/Administracion.aspx','_blank')
              }}
              >{strings.Link}</Link>
            </div>
          </div>
        }
      </div>
    );
  }
}
