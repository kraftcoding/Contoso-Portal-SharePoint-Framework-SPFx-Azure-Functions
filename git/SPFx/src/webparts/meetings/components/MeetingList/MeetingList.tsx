import * as React from 'react';
import * as strings from 'MeetingsWebPartStrings';
import styles from './MeetingList.module.scss';
import { IMeetingListprops, IMeetingListState } from './IMeetingList';
import { Input, Link, Tag } from '@fluentui/react-components';
import { EventStatus, IEvent } from '../../../../service/BackendServiceModels/EventModels';
import { getTermLabel, getTimeFormatted } from '../../../../utils/Utils';
import { NavLink } from 'react-router-dom';
import { groupBy } from '@microsoft/sp-lodash-subset';
import { entries } from 'lodash/fp';
import { Location16Regular, SearchRegular, VideoChat16Regular } from '@fluentui/react-icons';
import format from 'date-fns/format';
import es from 'date-fns/locale/es';
import ca from 'date-fns/locale/ca';
import eu from 'date-fns/locale/eu';
import gl from 'date-fns/locale/gl';

const locales: any = { 'es-ES': es, 'ca-ES': ca, 'eu-ES': eu, 'gl-ES': gl };
export default class MeetingList extends React.Component<IMeetingListprops, IMeetingListState> {

    OnlineTerm: string = '8f4e0e22-2c3f-4454-8210-f4151dceeb89';
    OnlineInPersonTerm: string = 'e49c8fe8-87f5-4b62-b6ce-0df240b7ec7f';
    AtendanceTypeSetId: string = '0775b014-e8fc-4aac-ae49-6b0b901156c1';
    OnlineToolSetId: string = '55dfa3e7-f0df-4dc8-b53c-4018418041fe';

    constructor(props: IMeetingListprops) {
        super(props);
        this.state = {
            nextEvents: [],
            pastEvents: [],
            archivedEvents: [],
            selectedEvents: [],
            selectedTab: undefined,
            isMobile: window.innerWidth <= 640
        };
        this.handleResize = this.handleResize.bind(this);
    }

    public componentDidMount(): void {
        window.addEventListener('resize', this.handleResize);
        this.onInit(undefined);
    }

    public componentWillUnmount(): void {
        window.removeEventListener('resize', this.handleResize);
    }

    public handleResize(): void {
        this.setState({ isMobile: window.innerWidth <= 640 });
    }

    private tabSelect(value: string): void {
        const { onSelectEvent } = this.props;
        const { nextEvents, pastEvents, archivedEvents } = this.state;

        let selectedEvents: IEvent[] = [];

        switch (value) {
            case "prodximas":
                selectedEvents = nextEvents;
                onSelectEvent(nextEvents[0]);
                break;
            case "celebradas":
                selectedEvents = pastEvents;
                onSelectEvent(pastEvents[0]);
                break;
            case "archivadas":
                selectedEvents = nextEvents;
                onSelectEvent(archivedEvents[0]);
                break;
        }

        this.setState({ selectedTab: value, selectedEvents });
    }

    public componentDidUpdate(testvprops: Readonly<IMeetingListprops>, testvState: Readonly<IMeetingListState>, snapshot?: any): void {
        if (testvprops.selectedEvent !== this.props.selectedEvent) {
            this.onInit(undefined);
        }
    }

    private onInit(tabToSelect?: string): void {
        const { onSelectEvent, match, events } = this.props;
        const { selectedTab } = this.state;

        let nextEvents: IEvent[] = [];
        let pastEvents: IEvent[] = [];
        let archivedEvents: IEvent[] = [];
        let selectedEvents: IEvent[] = [];

        events.forEach((item: IEvent): void => {
            switch (item.StatusId) {
                case EventStatus.InConstruction:
                case EventStatus.testBooking:
                case EventStatus.Published:
                case EventStatus.InCelebration:
                    nextEvents.push(item);
                    break;
                case EventStatus.Celebrated:
                case EventStatus.Canceled:
                    pastEvents.push(item);
                    break;
                case EventStatus.Archived:
                    archivedEvents.push(item);
                    break;
            }
        });

        if (events.length !== 0) {
            if (match.params.id) {
                if (match.params.id === "-1") {
                    if (nextEvents && nextEvents.length !== 0) {
                        onSelectEvent(nextEvents[0]);
                        tabToSelect = "prodximas";
                    }
                    else if (pastEvents && pastEvents.length !== 0) {
                        onSelectEvent(pastEvents[0]);
                        tabToSelect = "celebradas";
                    }
                    else if (archivedEvents && archivedEvents.length !== 0) {
                        onSelectEvent(archivedEvents[0]);
                        tabToSelect = "archivadas";
                    }
                }
                else {
                    const eventToSelect = events.find((event: IEvent) => event.Id === match.params.id);

                    if (!selectedTab) {
                        onSelectEvent(eventToSelect);
                    }
                    switch (eventToSelect?.StatusId) {
                        case EventStatus.InConstruction:
                        case EventStatus.testBooking:
                        case EventStatus.Published:
                        case EventStatus.InCelebration:
                            tabToSelect = "prodximas";
                            selectedEvents = nextEvents;
                            break;
                        case EventStatus.Celebrated:
                        case EventStatus.Canceled:
                            tabToSelect = "celebradas";
                            selectedEvents = pastEvents;
                            break;
                        case EventStatus.Archived:
                            tabToSelect = "archivadas";
                            selectedEvents = archivedEvents;
                            break;
                    }
                }
            }
        }

        this.setState({
            pastEvents,
            nextEvents,
            archivedEvents,
            selectedEvents,
            selectedTab: tabToSelect
        });
    }

    public onSearch(ev: React.ChangeEvent<HTMLInputElement>): void {
        const { nextEvents, pastEvents, archivedEvents, selectedTab } = this.state;

        let eventsToFilter: IEvent[];

        switch (selectedTab) {
            case "prodximas":
                eventsToFilter = nextEvents;
                break;
            case "celebradas":
                eventsToFilter = pastEvents;
                break;
            case "archivadas":
                eventsToFilter = archivedEvents;
                break;
            default:
                eventsToFilter = [];
        }

        if (ev.target.value) {
            this.setState({ selectedEvents: eventsToFilter.filter(event => event.Title.toLocaleLowerCase().includes(ev.target.value.toLocaleLowerCase())) });
        } else {
            this.setState({ selectedEvents: eventsToFilter });
        }
    }

    public render(): React.ReactElement<IMeetingListprops> {
        const { context, onSelectEvent, statusTerms, selectedEvent, loadingEventId, showArchived } = this.props;
        const { nextEvents, pastEvents, archivedEvents, selectedTab, selectedEvents, isMobile } = this.state;
        const locale = context.pageContext.cultureInfo.currentUICultureName;
        const sorted = groupBy(selectedEvents, (event: IEvent) => {
            const eventStart: Date = new Date(event.StartDate);
            eventStart.setHours(0, 0, 0);
            return eventStart;
        });

        return (
            <section className={styles.meetingsList}>
                <div className={styles.tabNavLinks} key={selectedEvent?.Id + "meetinglist"}>
                    {
                        /* Pestaña "Próximas" */
                        (nextEvents.length !== 0) &&
                        <NavLink
                            activeClassName={styles.active}
                            onClick={() => this.tabSelect("prodximas")}
                            to={`/${selectedTab === "prodximas" ? selectedEvent?.Id : nextEvents[0].Id}`}
                        >
                            <div className={styles.linkItem}>
                                {`${strings.NextMeetings} (${nextEvents.length})`}
                            </div>
                        </NavLink>
                    }
                    {
                        /* Pestaña "Celebradas" */
                        (pastEvents.length !== 0) &&
                        <NavLink
                            activeClassName={styles.active}
                            onClick={() => this.tabSelect("celebradas")}
                            to={`/${selectedTab === "celebradas" ? selectedEvent?.Id : pastEvents[0].Id}`}
                        >
                            <div className={styles.linkItem}>
                                {`${strings.CelebratedMeetings} (${pastEvents.length})`}
                            </div>
                        </NavLink>
                    }
                    {
                        /* Pestaña "Archivadas" */
                        (archivedEvents.length !== 0 && showArchived) &&
                        <NavLink
                            activeClassName={styles.active}
                            onClick={() => this.tabSelect("archivadas")}
                            to={`/${selectedTab === "archivadas" ? selectedEvent?.Id : archivedEvents[0].Id}`}
                        >
                            <div className={styles.linkItem}>
                                {`${strings.ArchivedMeetings} (${archivedEvents.length})`}
                            </div>
                        </NavLink>
                    }
                </div>
                <div className={styles.searchContainer}>
                    <div>
                        <Input
                            type='text'
                            aria-label={strings.LookFor}
                            placeholder={strings.LookFor}
                            onChange={this.onSearch.bind(this)}
                            contentBefore={<SearchRegular />}
                        />
                    </div>
                </div>
                <div className={styles.meetingsListContainer}>

                    {selectedEvent && selectedTab && sorted && entries(sorted).map((events: [string, IEvent[]]) => {
                        const date = new Date(events[0]);
                        const convoDate: string = format(date, 'PPPP', { locale: locales[locale] });
                        const convoDateCap: string = convoDate && convoDate.charAt(0).toUpperCase() + convoDate.slice(1);

                        return (
                            <>
                                <div className={styles.dateGroup} key={loadingEventId + "list"}>{convoDateCap}</div>
                                {events[1].map((event: IEvent, i: number) => {
                                    const cssStatusClass = {
                                        [EventStatus.Celebrated]: styles.Celebrated,
                                        [EventStatus.InCelebration]: styles.InCelebration,
                                        [EventStatus.InConstruction]: styles.InConstruction,
                                        [EventStatus.testBooking]: styles.testBooking,
                                        [EventStatus.Published]: styles.Published,
                                        [EventStatus.Canceled]: styles.Canceled,
                                        [EventStatus.Archived]: styles.Archived
                                    }[event.StatusId];

                                    return (
                                        <NavLink
                                            className={styles.meetingsListItem}
                                            activeClassName={styles.listItemSelected}
                                            to={`/${event.Id}/`}
                                            onClick={() => onSelectEvent(event)}
                                            ref={(node: HTMLAnchorElement | null): void => {
                                                const isResponsive: boolean = window.innerWidth <= 640;
                                                if (node && !isResponsive && selectedEvent?.Id === event.Id) {
                                                    node.scrollIntoView({
                                                        behavior: "smooth",
                                                        block: "nearest"
                                                    });
                                                }
                                            }}
                                        >
                                            <div className={styles.datesAndStateRow}>
                                                <span className={styles.dates}>{getTimeFormatted(event.StartDate)} - {getTimeFormatted(event.EndDate)}</span>
                                                <span className={styles.state}>
                                                    <Tag className={`${styles.eventStatus} ${cssStatusClass}`} size='extra-small' appearance="brand" shape='circular'>
                                                        {getTermLabel(statusTerms, event.StatusId)}
                                                    </Tag>
                                                </span>
                                            </div>

                                            {
                                                ((event.AttendanceTypeId === this.OnlineTerm || event.AttendanceTypeId === this.OnlineInPersonTerm) &&


                                                    event.StatusId === EventStatus.InCelebration ?

                                                    <div className={styles.titleAndButtonRow}>
                                                        <span className={styles.title} style={{ flex: 1 }} title={event.Title}>{event.Title}</span>
                                                        <div className={styles.buttonJoinContainer}>
                                                            {isMobile && event.MeetingToolUrl &&
                                                                <Link
                                                                    className={styles.buttonJoin}
                                                                    href={event.MeetingToolUrl}
                                                                    target="_blank"
                                                                    onClick={(e) => e.stopprodpagation()}
                                                                >
                                                                    <VideoChat16Regular /> {strings.VideoconferenceMeetingList}
                                                                </Link>
                                                            }
                                                        </div>
                                                    </div>
                                                    :
                                                    <span className={styles.title} title={event.Title}>{event.Title}</span>
                                                )
                                            }

                                            {event.Location &&
                                                <div className={styles.locationState}>
                                                    <Location16Regular />
                                                    <span>{event.Location}</span>
                                                </div>
                                            }
                                        </NavLink>
                                    );
                                })}
                            </>
                        );
                    })}
                </div>
            </section>
        );
    }
}