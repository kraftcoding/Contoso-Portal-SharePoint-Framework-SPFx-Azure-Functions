import * as React from 'react';
import * as strings from 'MyCollaborationSpacesWebPartStrings';
import styles from './MyCollaborationSpaces.module.scss';
import type { IMyCollaborationSpacesprops } from './IMyCollaborationSpacesprops';
import type { IMyCollaborationSpacesState } from './IMyCollaborationSpacesState';
import { Teams } from '../../../models/ITeams';
import { Card, CardHeader, Skeleton, SkeletonItem } from '@fluentui/react-components';
import { TeamsIcon } from '@fluentui/react-icons-northstar';

export default class MyCollaborationSpaces extends React.Component<IMyCollaborationSpacesprops, IMyCollaborationSpacesState> {

  constructor(props: IMyCollaborationSpacesprops) {
    super(props);
    this.state = {
      teams: [],
      loading: false,
    }
  }

  public componentDidMount(): void {
    void this.onInit();
  }

  public shouldComponentUpdate(nextprops: Readonly<IMyCollaborationSpacesprops>, nextState: Readonly<IMyCollaborationSpacesState>): boolean {
    return nextprops !== this.props || nextState !== this.state;
  }

  private async onInit(): promise<void> {
    const { spService } = this.props;
    this.setState({ loading: true });
    const teams: Teams[] = await spService.getMyTeams();
    const channels = await promise.all(teams.map(team => spService.getTeamsPrimaryChannel(team.id)));
   // const informationTeams: Teams[] = await promise.all(teams.map(team => spService.getTeamsInformation(team.id)));
   // const teamsUrls: string[] = informationTeams.map(team => team.webUrl);
    teams.forEach((team, i) => {
      //team.webUrl = teamsUrls[i];
      //$url= "https://teams.microsoft.com/v2/?tenantId="+$TenantId+"#/l/team/"+$channelId+"/conversations?groupId="+$groupId+"&ngc=true"
      const tenantId = team.tenantId;
      const teamId = team.id;
      const channelId = channels[i].id;
      const url = `https://teams.microsoft.com/v2/?tenantId=${tenantId}#/l/team/${channelId}/conversations?groupId=${teamId}`;
      team.webUrl = url;
    });
    this.setState({ teams, loading: false });
    await this.initPictures();
  }

  private async initPictures(): promise<void> {
    const { spService } = this.props;
    const teams: Teams[] = this.state.teams;
    const blobs: Blob[] = await promise.all(this.state.teams.map(team => spService.getTeamsPhoto(team.id)));
    const teamsPhotos: string[] = blobs.map(blob => window.URL.createObjectURL(blob));
    teams.forEach((team, i) => {
      team.photo = teamsPhotos[i];
    });
    this.setState({ teams });
  }

  public renderMoreLink = (): React.ReactElement => {
    return <a className={styles.moreLink} href={this.props.moreLink} title={strings.ViewMoreCollaborationSpaces}>{strings.MoreLink}</a>;
  };

  public render(): React.ReactElement<IMyCollaborationSpacesprops> {
    const { title, showDescription, moreLink } = this.props;
    return (
      <section className={styles.myCollaborationSpaces}>
        <div className={styles.titleRow}>
          <div className={styles.wpTitle}>{title}</div>
          {
            moreLink && this.renderMoreLink()
          }
        </div>
        <div className={styles.mainContainer}>
          {
            showDescription &&
            <div className={styles.descriptionContainer}>
              <div className={styles.logoContainer}> <TeamsIcon className={styles.teamsLogo} /> </div>
              <span className={styles.mainText}> {strings.MainContainerText} </span>
            </div>
          }
          <div className={styles.cardsContainer}>
            {this.state.loading ?
              <>
                <Skeleton className={styles.skeletonCard}>
                  <SkeletonItem style={{ borderRadius: '4px' }} shape="square" size={48} />
                  <SkeletonItem shape="rectangle" size={20} />
                </Skeleton>
                <Skeleton className={styles.skeletonCard}>
                  <SkeletonItem style={{ borderRadius: '4px' }} shape="square" size={48} />
                  <SkeletonItem shape="rectangle" size={20} />
                </Skeleton>
                <Skeleton className={styles.skeletonCard}>
                  <SkeletonItem style={{ borderRadius: '4px' }} shape="square" size={48} />
                  <SkeletonItem shape="rectangle" size={20} />
                </Skeleton>
              </>
              :
              this.state.teams.map((team, i) => {
                return (
                  <Card key={i} className={styles.card} appearance='subtle'>
                    <CardHeader
                      image={
                        !team.photo ?
                          <Skeleton>
                            <SkeletonItem style={{ borderRadius: '4px' }} shape="square" size={48} />
                          </Skeleton>
                          :
                          <img alt={`logotipo ${team.displayName}`} className={styles.cardImage} src={team.photo} />
                      }
                      header={<a key={i} href={team.webUrl} className={styles.linkSee} target='_blank' rel="noreferrer" >
                        <span className={styles.cardText}>{team.displayName}</span>
                      </a>}
                    />
                  </Card>
                );
              })
            }
          </div>
        </div>
      </section>
    );
  }

}