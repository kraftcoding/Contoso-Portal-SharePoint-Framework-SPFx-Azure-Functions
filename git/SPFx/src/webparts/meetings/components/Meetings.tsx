import * as React from 'react';
import styles from './Meetings.module.scss';
import { IMeetingsprops, IMeetingsState } from './IMeetings';
import { Switch, Route, HashRouter, NavLink, Redirect } from "react-router-dom";
import MeetingList from './MeetingList/MeetingList';
import MeetingStatusBar from './MeetingStatusBar/MeetingStatusBar';
import MeetingPoints from './MeetingPoints/MeetingPoints';
import MeetingAssistants from './MeetingAssistants/MeetingAssistants';
import MeetingDocs from './MeetingDocs/MeetingDocs';
import MeetingMinutes from './MeetingMinutes/MeetingMinutes';
import { Divider, Image, Spinner } from '@fluentui/react-components';
import { Calendar16Regular, Folder16Regular, Handshake16Regular, Notepad16Regular, PeopleCommunity16Regular, VoteRegular } from '@fluentui/react-icons';
import { EventStatus, IEvent } from '../../../service/BackendServiceModels/EventModels';
import * as strings from 'MeetingsWebPartStrings';
import { Logger } from '../../../utils/Logger';
import { ITermInfo } from '@pnp/sp/taxonomy/types';
import { Term } from '../../../models/ITag';
import MeetingAgreements from './MeetingAgreements/MeetingAgreements';
import { BodyRole, RoleGroups, userIsAnyOf, userIsRoleFromGroupName } from '../../../service/RoleService';
import Calls from './MeetingNew/Calls';
import MeetingVoting from './MeetingVoting/MeetingVoting';

const statusSetId: string = '5e7c267a-4f4f-4d43-99c8-a94d1660ae85';

export default class Meetings extends React.Component<IMeetingsprops, IMeetingsState> {

  constructor(props: IMeetingsprops) {
    super(props);
    this.state = {
      events: [],
      selectedEvent: undefined,
      isLoading: true,
      statusTerms: [],
      loadingChangedEvent: false,
      loadingChangedEventId: undefined,
      isEditor: false,
      isGuest: false,
      selectedEventReadOnly: false,
      showArchived:false,
      isOCprodle: false
    };
  }

  public componentDidMount(): void {
    void this.onInit();
  }

  private async onInit(): promise<void> {
    const { spService, bkService, locale, context } = this.props;

    try {
      this.setState({ isLoading: true });

      let [statusRes, eventsRes, isEditor, isGuest, showArchived]: [ITermInfo[], IEvent[], boolean, boolean, boolean] = await promise.all([
        await spService.getTaxonomy(statusSetId),
        await bkService.getEvents(context.pageContext.web.serverRelativeUrl.split("sites/")[1]),
        userIsAnyOf(context, spService, [BodyRole.Scheduler, BodyRole.SchedulerAssistant]),
        userIsAnyOf(context, spService, [BodyRole.Guest]),
        userIsAnyOf(context, spService, [BodyRole.Scheduler, BodyRole.SchedulerAssistant, BodyRole.Member, BodyRole.MemberAssistant]),
      ]);
      var isOCprodle = false;
      if(!showArchived){
        const isOCP= await userIsRoleFromGroupName(spService, RoleGroups.OficinaConferenciatestsidentes);
        showArchived = isOCP;
        isOCprodle = isOCP && !isGuest;
      }
      if (showArchived) {
        const archivedEvents = await bkService.getEvents(context.pageContext.web.serverRelativeUrl.split("sites/")[1], true);
        eventsRes = eventsRes.filter(event => event.StatusId !== EventStatus.Archived);
        eventsRes.push(...archivedEvents);
      }

      const statusTerms: Term[] = statusRes.map(res => Term.mapSPToTerm(res, locale));
      const events: IEvent[] = eventsRes.map(event => {
        return {
          ...event,
          StartDate: new Date(event.StartDate.toString()),
          EndDate: new Date(event.EndDate.toString())
        }
      }).sort((a: IEvent, b: IEvent) => {
        return a.StartDate.getTime() - b.StartDate.getTime();
      });

      this.setState({ isLoading: false, events, statusTerms, isEditor, isGuest, showArchived, isOCprodle });
    }
    catch (error) {
      Logger.error("Error OnInit - Meetings", error, context);
      this.setState({ isLoading: false });
    }
  }

  private onSelectEvent = (event: IEvent): void => {
    if (event) {
      const readOnly = event?.StatusId === EventStatus.Canceled || event?.StatusId === EventStatus.Archived;
      this.setState({
        selectedEvent: event, selectedEventReadOnly: readOnly
      });
    }
  }

  private onChangedEvent = (event: IEvent): void => {
    this.setState({ loadingChangedEvent: true, loadingChangedEventId: event.Id });
    this.changeEventInState(event);
  }

  private changeEventInState(event: IEvent): void {
    let events = [...this.state.events];

    event = {
      ...event,
      StartDate: new Date(event.StartDate),
      EndDate: new Date(event.EndDate)
    }

    for (let i = 0; i < events.length; i++) {
      if (events[i].Id === event.Id) {
        events[i] = event;
        break;
      }
    }

    this.setState({ events: events, selectedEvent: event, loadingChangedEvent: false, loadingChangedEventId: undefined });
  }

  private statusChange = async (nextStatusId: string): promise<void> => {
    const { bkService } = this.props;
    const { selectedEvent } = this.state;

    this.setState({ loadingChangedEvent: true, loadingChangedEventId: selectedEvent?.Id });

    try {
      const result = selectedEvent && await bkService.changeEventStatusById(selectedEvent?.BodyId, selectedEvent?.Id, nextStatusId, "");

      if (result) {
        const refreshedEvent: IEvent = await bkService.getEvent(result.BodyId, result.SharedEventId);
        this.changeEventInState(refreshedEvent);
      }
    }
    catch (error) {
      Logger.error("Error al cambiar el estado", error, this.context);
      this.setState({ loadingChangedEvent: false, loadingChangedEventId: undefined });
    }
  }

  public render(): React.ReactElement<IMeetingsprops> {
    const { title, showVoting } = this.props;
    const {
      events,
      isLoading,
      selectedEvent,
      statusTerms,
      loadingChangedEvent,
      loadingChangedEventId,
      isEditor,
      isGuest,
      selectedEventReadOnly,
      showArchived,
      isOCprodle
    } = this.state;

    const isArchiving = (selectedEvent?.StatusId === EventStatus.Archived &&
      (selectedEvent?.StorageServerRelativeUrl.indexOf("MeetingsConstruccion") != -1 || selectedEvent?.StorageServerRelativeUrl.indexOf("MeetingsFinales") != -1)) ? true : false;

    return (
      <>
        <div className={styles.itemRow}>
          <div className={styles.itemColumnGrow}>
            <div className={styles.wpTitle}>{title}</div>
          </div>
          <div className={styles.itemColumn}>
            {isEditor && <Calls {...this.props}></Calls>}
          </div>
        </div>
        <section className={styles.meetings}>
          {
            <HashRouter>
              { // Router por defecto
                (!isLoading && events && !selectedEvent) &&
                <Route exact path="/" render={(props) => (
                  <Redirect to={`/-1`} />
                )} />
              }
              <Route path={"/:id"} render={(parentprops: any) =>
                <div className={styles.mainContainer}>
                  <div className={styles.list}>
                    {
                      isLoading ?
                        <Spinner style={{ flex: 1 }} label={`${strings.Loading}...`} />
                        :
                        (
                          (events.length !== 0) ?
                            <MeetingList
                              {...parentprops}
                              {...this.props}
                              events={[...events]}
                              statusTerms={statusTerms}
                              onSelectEvent={this.onSelectEvent}
                              selectedEvent={{ ...selectedEvent }}
                              loadingStatus={loadingChangedEvent}
                              loadingEventId={loadingChangedEventId}
                              key={"meetinglist"}
                              isEditor={isEditor}
                              showArchived={showArchived}
                            />
                            :
                            <div className={styles.noElementsContainer}>
                              <span> {strings.ThereAreNoMeetingsToShow} </span>
                            </div>
                        )
                    }
                  </div>
                  <Divider className={styles.dividerVertical} vertical />
                  <div className={styles.convoContainer}>
                    {
                      selectedEvent &&
                      <MeetingStatusBar
                        {...this.props}
                        isEditor={isEditor}
                        onDataChanged={this.onChangedEvent.bind(this)}
                        onStatusChange={this.statusChange}
                        loadingStatus={loadingChangedEvent}
                        loadingEventId={loadingChangedEventId}
                        event={selectedEvent}
                        statusTerms={statusTerms}
                        readOnly={selectedEventReadOnly}
                        key={selectedEvent?.Id + "statusbar" + loadingChangedEventId}
                        isOCprodle={isOCprodle}
                      />
                    }
                    {
                      (events.length !== 0) &&
                      <Divider className={styles.dividerHorizontal} />
                    }
                    {
                      (!isLoading && events.length === 0) &&
                      <div className={styles.noElementsContainer}>
                        <div><Image fit='default' src={require('../assets/welcome-light.png')} ></Image></div>
                      </div>
                    }
                    {
                      (!isLoading && selectedEvent && isArchiving) &&
                      <div className={styles.noElementsContainer}>
                        <div> {strings.TheMeetingIsBeingArchived} </div>
                      </div>
                    }
                    {
                      (!isLoading && !selectedEvent) &&
                      <div className={styles.noElementsContainer}>
                        <div> {strings.TheMeetingIsNotAvailable} </div>
                      </div>
                    }
                    {
                      (selectedEvent && !isArchiving && events.length !== 0) &&
                      <div className={styles.router}>
                        <div
                          className={styles.tabNavLinks}
                          ref={(node: HTMLDivElement | null): void => {
                            const isResponsive: boolean = window.innerWidth <= 640;
                            if (node && isResponsive) {
                              node.scrollIntoView({
                                behavior: "smooth",
                                block: "nearest"
                              });
                            }
                          }}
                        >
                          {/* Pestaña "Orden del día" */}
                          <NavLink exact activeClassName={styles.active} to={`/${selectedEvent.Id}/ordenes`}>
                            <div className={styles.linkItem}>
                              <Calendar16Regular className={styles.icon16} /><span>{strings.MeetingPoints}</span>
                            </div>
                          </NavLink>
                          {/* Pestaña "Asistentes" */}
                          <NavLink exact activeClassName={styles.active} to={`/${selectedEvent.Id}/asistentes`}>
                            <div className={styles.linkItem}>
                              <PeopleCommunity16Regular className={styles.icon16} /><span>{strings.Assistants}</span>
                            </div>
                          </NavLink>
                          {/* Pestaña "Documents */}
                          <NavLink exact activeClassName={styles.active} to={`/${selectedEvent.Id}/docs`}>
                            <div className={styles.linkItem}>
                              <Folder16Regular className={styles.icon16} /><span>{strings.Documents}</span>
                            </div>
                          </NavLink>
                          { /* Pestaña "Votación" */
                            showVoting && (
                              (selectedEvent.StatusId === EventStatus.InCelebration) ||
                              (selectedEvent.StatusId === EventStatus.Celebrated && isEditor) ||
                              (selectedEvent.StatusId === EventStatus.Archived && isEditor) ||
                              (selectedEvent.StatusId === EventStatus.Canceled && isEditor)
                            ) &&
                            <NavLink activeClassName={styles.active} to={`/${selectedEvent.Id}/Vote`}>
                              <div className={styles.linkItem}>
                                <VoteRegular className={styles.icon16} /> <span> {"Votación"} </span>
                              </div>
                            </NavLink>
                          }
                          {
                            /* Pestaña "Acuerdos" */
                            (
                              (selectedEvent.StatusId === EventStatus.InCelebration && isEditor) ||
                              (selectedEvent.StatusId === EventStatus.InCelebration && !showVoting) ||
                              (selectedEvent.StatusId === EventStatus.Celebrated) ||
                              (selectedEvent.StatusId === EventStatus.Archived) ||
                              (selectedEvent.StatusId === EventStatus.Canceled)
                            ) &&
                            <NavLink exact activeClassName={styles.active} to={`/${selectedEvent.Id}/acuerdos`}>
                              <div className={styles.linkItem}>
                                <Handshake16Regular className={styles.icon16} /><span>{strings.Agreements}</span>
                              </div>
                            </NavLink>}
                          {
                            /* Pestaña "Minutes" */
                            (
                              (selectedEvent.StatusId !== EventStatus.InConstruction) &&
                              (selectedEvent.StatusId !== EventStatus.testBooking) &&
                              (selectedEvent.StatusId !== EventStatus.Published) &&
                              (selectedEvent.StatusId !== EventStatus.InCelebration)
                            ) &&
                            <NavLink exact activeClassName={styles.active} to={`/${selectedEvent.Id}/Minutes`}>
                              <div className={styles.linkItem}>
                                <Notepad16Regular className={styles.icon16} /><span>{strings.Minutes}</span>
                              </div>
                            </NavLink>}
                        </div>
                        <div className={styles.tabsWrapper}>
                          <Switch>
                            <Route exact path="/:id" render={(props) => (
                              <Redirect to={`/${selectedEvent.Id}/ordenes`} />
                            )} />
                            <Route
                              exact
                              path='/:id/asistentes'
                              render={(props: any) => <MeetingAssistants key={selectedEvent?.Id + "asistentes" + loadingChangedEventId}
                                isEditor={isEditor} readOnly={selectedEventReadOnly} event={selectedEvent} {...props} {...this.props} />}
                            />
                            <Route
                              exact
                              path='/:id/ordenes'
                              render={(props: any) => <MeetingPoints key={selectedEvent?.Id + "ordenes" + loadingChangedEventId}
                                isEditor={isEditor} readOnly={selectedEventReadOnly} event={selectedEvent} {...props} {...this.props} />}
                            />
                            <Route
                              exact
                              path='/:id/docs'
                              render={(props: any) => <MeetingDocs key={selectedEvent?.Id + "docs" + loadingChangedEventId}
                                isEditor={isEditor} readOnly={selectedEventReadOnly} event={selectedEvent} {...props} {...this.props} />}
                            />
                            <Route
                              exact
                              // path='/:id/Vote'
                              path='/:id/Vote/:optionalParam?'
                              render={(props: any) => <MeetingVoting key={selectedEvent?.Id + "Vote" + loadingChangedEventId}
                                isEditor={isEditor} readOnly={selectedEventReadOnly} event={selectedEvent} {...props} {...this.props} />}
                            />
                            <Route
                              exact
                              path='/:id/acuerdos'
                              render={(props: any) => <MeetingAgreements key={selectedEvent?.Id + "acuerdos" + loadingChangedEventId}
                                isEditor={isEditor} isGuest={isGuest} isOCprodle={isOCprodle} readOnly={selectedEventReadOnly} event={selectedEvent} {...props} {...this.props} />}
                            />
                            <Route
                              exact
                              path='/:id/Minutes'
                              render={(props: any) => <MeetingMinutes key={selectedEvent?.Id + "Minutes" + loadingChangedEventId}
                                event={selectedEvent} readOnly={selectedEventReadOnly} isEditor={isEditor} {...props} {...this.props} />}
                            />
                          </Switch>
                        </div>
                      </div>
                    }
                  </div>
                </div>}
              />
            </HashRouter>
          }
        </section>
      </>
    );
  }

}