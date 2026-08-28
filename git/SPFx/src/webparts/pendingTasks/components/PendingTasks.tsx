import * as React from 'react';
import * as strings from 'PendingTasksWebPartStrings';
import styles from './PendingTasks.module.scss';
import {
  IPendingTasksprops,
  IPendingTasksState
} from './IPendingTasks';
import {
  IUserBodyRole,
  userBodyRoles
} from '../../../service/RoleService';
import { IUserTaskModel } from '../../../service/BackendServiceModels/UserTaskModel';
import { Spinner } from '@fluentui/react-components';
import { ITermInfo } from '@pnp/sp/taxonomy';
import { Term } from '../../../models/ITag';
import MeetingAttendanceRequest from './PendingTasksDialogs/MeetingAttendanceRequest';
import DelegatedMeetingAttendanceRequest from './PendingTasksDialogs/DelegatedMeetingAttendanceRequest';
import MinutesApprodvalRequest from './PendingTasksDialogs/MinutesApprodvalRequest';
import MinutesModificationRequest from './PendingTasksDialogs/MinutesModificationRequest';
import CertificationRequest from './PendingTasksDialogs/CertificationRequest';

const OrgansSetId: string = "6f98455a-a1c9-45e7-8c5d-aaca87690eff";

export enum PendingTaskType {
  /* Request de Attendance a Meeting */
  MeetingAttendanceRequest = "0x0101006B46B0FAC7593A4E8EFEA5A5399F44AC0102010301",
  /* Request de delegación de Attendance a Meeting */
  DelegatedMeetingAttendanceRequest = "0x0101006B46B0FAC7593A4E8EFEA5A5399F44AC0102010302",
  /* Request de aprodbación de Minutes */
  MinutesApprodvalRequest = "0x0101006B46B0FAC7593A4E8EFEA5A5399F44AC0102010305",
  /* Request de modificación de Minutes */
  MinutesModificationRequest = "0x0101006B46B0FAC7593A4E8EFEA5A5399F44AC0102010304",
  /* Request de certificación */
  CertificationRequest = "0x0101006B46B0FAC7593A4E8EFEA5A5399F44AC0102010303"
}

export default class PendingTasks extends React.Component<IPendingTasksprops, IPendingTasksState> {

  private organTaxonomyTerms: Term[] = [];

  constructor(props: IPendingTasksprops) {
    super(props);
    this.state = {
      loadingThePendingTasks: true,
      pendingTasks: []
    };
  }

  public async componentDidMount(): promise<void> {
    await this.getTheOrganTaxonomyTerms();
    await this.getThePendingTasks();
  }

  public async componentDidUpdate(testvprops: Readonly<IPendingTasksprops>): promise<void> {
    const {
      numberOfResults,
      getThePendingTasksFromThisSite
    } = this.props;

    if (
      (numberOfResults !== testvprops.numberOfResults) ||
      (getThePendingTasksFromThisSite !== testvprops.getThePendingTasksFromThisSite)
    ) {
      await this.getTheOrganTaxonomyTerms();
      await this.getThePendingTasks();
    }
  }

  public async getThePendingTasks(): promise<void> {
    const {
      context,
      bkService,
      spService,
      numberOfResults,
      getThePendingTasksFromThisSite
    } = this.props;

    const bodyId: string = context.pageContext.web.serverRelativeUrl.split("sites/")[1];

    this.setState({
      loadingThePendingTasks: true
    });
    try {
      let pendingTasks: IUserTaskModel[] = [];

      if (getThePendingTasksFromThisSite) {
        pendingTasks = await bkService.getUserTaskByBodyId(bodyId);
      }
      else {
        const userBoDyRoles: IUserBodyRole[] = await userBodyRoles(spService);
        const allBodiespromises: promise<void | IUserTaskModel[]>[] = [];

        userBoDyRoles.forEach((body: IUserBodyRole): void => {
          allBodiespromises.push(bkService.getUserTaskByBodyId(body?.bodyId).then((res: IUserTaskModel[]): void => {
            pendingTasks.push(...res);
          }).catch(error => console.log(error)));
        });
        await promise.all(allBodiespromises);
      }

      pendingTasks.sort((a: IUserTaskModel, b: IUserTaskModel): number =>
        new Date(a.TaskEndDate).getTime() - new Date(b.TaskEndDate).getTime());

      const slicedPendingTasks: IUserTaskModel[] = (
        (getThePendingTasksFromThisSite) ?
          pendingTasks
          :
          pendingTasks.slice(0, numberOfResults)
      );

      const filteredPendingTasks: IUserTaskModel[] = slicedPendingTasks.filter((pendingTask: IUserTaskModel): boolean => {
        const taskType: string = pendingTask.TaskTypeId.slice(0, 50);
        return [
          PendingTaskType.MeetingAttendanceRequest.toString(),
          PendingTaskType.DelegatedMeetingAttendanceRequest.toString(),
          PendingTaskType.MinutesApprodvalRequest.toString(),
          PendingTaskType.MinutesModificationRequest.toString(),
          PendingTaskType.CertificationRequest.toString()
        ].includes(taskType);
      });

      this.setState({
        loadingThePendingTasks: false,
        pendingTasks: filteredPendingTasks
      });
    }
    catch (error) {
      console.error(error);

      this.setState({
        loadingThePendingTasks: false,
        pendingTasks: []
      });
    }
  }

  private async getTheOrganTaxonomyTerms(): promise<void> {
    const {
      spService,
      locale,
      getThePendingTasksFromThisSite
    } = this.props;

    if (!getThePendingTasksFromThisSite && (this.organTaxonomyTerms.length === 0)) {
      try {
        const organTaxonomy: ITermInfo[] = await spService.getTaxonomy(OrgansSetId);
        this.organTaxonomyTerms = organTaxonomy?.map((tag: ITermInfo): Term => Term.mapSPToTerm(tag, locale));
      }
      catch (error) {
        console.error(error);
      }
    }
  }

  private getTheDepartmentName(bodyNameId: string): string {
    const selectedOrganTaxonomyTerm: Term | undefined = this.organTaxonomyTerms?.find(
      (organTaxonomyTerm: Term): boolean => organTaxonomyTerm.key === bodyNameId);
    const DepartmentName: string = selectedOrganTaxonomyTerm?.text || "";

    return DepartmentName;
  }

  public renderMoreLink = (): React.ReactElement => {
    return (
      <a className={styles.moreLink} href={this.props.moreLink} title={strings.ViewMorePendingTasks}>
        {strings.MoreLink}
      </a>
    );
  };

  public render(): React.ReactElement<IPendingTasksprops> {
    const {
      context,
      spService,
      bkService,
      locale,
      title,
      moreLink,
      getThePendingTasksFromThisSite
    } = this.props;
    const {
      loadingThePendingTasks,
      pendingTasks
    } = this.state;

    return (
      <section className={styles.pendingTasks}>
        <div className={styles.titleRow}>
          <div className={styles.wpTitle}>
            {title}
          </div>
          {
            (moreLink) &&
            this.renderMoreLink()
          }
        </div>
        <div className={styles.mainContainer}>
          {
            (loadingThePendingTasks) ?
              <div className={styles.messageContainer}>
                <Spinner style={{ flex: 1 }} label={strings.LoadingThePendingTasks + "..."} />
              </div>
              :
              (pendingTasks && (pendingTasks.length > 0)) ?
                pendingTasks.map((pendingTask: IUserTaskModel, index: number): JSX.Element => {
                  const commonDialogproperties = {
                    locale,
                    context: context,
                    spService: spService,
                    bkService,
                    organTaxonomyTerms: this.organTaxonomyTerms,
                    pendingTask,
                    getThePendingTasks: (): promise<void> => this.getThePendingTasks(),
                    getThePendingTasksFromThisSite,
                    DepartmentName: this.getTheDepartmentName(pendingTask.BodyNameId)
                  };

                  return (
                    <>
                      {
                        {
                          /* Request de Attendance a Meeting */
                          [PendingTaskType.MeetingAttendanceRequest]:
                            <MeetingAttendanceRequest {...commonDialogproperties} />,
                          /* Request de delegación de Attendance a Meeting */
                          [PendingTaskType.DelegatedMeetingAttendanceRequest]:
                            <DelegatedMeetingAttendanceRequest {...commonDialogproperties} />,
                          /* Request de aprodbación de Minutes */
                          [PendingTaskType.MinutesApprodvalRequest]:
                            <MinutesApprodvalRequest {...commonDialogproperties} />,
                          /* Request de modificación de Minutes */
                          [PendingTaskType.MinutesModificationRequest]:
                            <MinutesModificationRequest {...commonDialogproperties} />,
                          /* Request de certificación */
                          [PendingTaskType.CertificationRequest]:
                            <CertificationRequest {...commonDialogproperties} />,
                        }[pendingTask.TaskTypeId.slice(0, 50)]
                      }
                      {
                        (index < (pendingTasks.length - 1)) &&
                        <div className={styles.spacer} />
                      }
                    </>
                  );
                })
                :
                <div className={styles.messageContainer}>
                  <span className={styles.noPendingTasksAvailable}> {strings.NoPendingTasksToShow} </span>
                </div>
          }
        </div>
      </section>
    );
  }

}