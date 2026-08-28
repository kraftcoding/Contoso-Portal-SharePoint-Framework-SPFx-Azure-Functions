/* eslint-disable @typescript-eslint/ban-ts-comment */
import * as React from 'react';
import styles from './AgendaAnnouncement.module.scss';
import type { IAgendaAnnouncementprops } from './IAgendaAnnouncementprops';
import { Calendar, dateFnsLocalizer, Navigate as navigate } from 'react-big-calendar';
import 'react-big-calendar/lib/css/react-big-calendar.css';
import format from 'date-fns/format';
import parse from 'date-fns/parse';
import startOfWeek from 'date-fns/startOfWeek';
import getDay from 'date-fns/getDay';
import { IAgendaAnnouncementState } from './IAgendaAnnouncementState';
import * as strings from 'AgendaAnnouncementWebPartStrings';
import ca from 'date-fns/locale/ca';
import eu from 'date-fns/locale/eu';
import gl from 'date-fns/locale/gl';
import es from 'date-fns/locale/es';
import { Popover, PopoverSurface, PopoverTrigger, Spinner } from '@fluentui/react-components';
import {
  ClockAlarm16Regular, Eye16Regular, PeopleTeam16Regular,
  ChevronLeft20Filled,
  ChevronRight20Filled
} from '@fluentui/react-icons';
import { IEvent, IEventFromBbdd } from '../../../service/BackendServiceModels/EventModels';
import { ITermInfo } from '@pnp/sp/taxonomy';
import { Term } from '../../../models/ITag';
import { IEventAgenda } from '../../../models/IEventAgenda';
import { getTermLabel, getTimeFormatted } from '../../../utils/Utils';
import { RoleGroups, userIsRoleFromGroupName } from '../../../service/RoleService';

const bodyTypeSetId: string = 'ba3ff2ef-09c8-4345-9bad-777014ddf636';
const bodyNameSetId: string = '6f98455a-a1c9-45e7-8c5d-aaca87690eff';
const attendancesFormatSetId: string = '0775b014-e8fc-4aac-ae49-6b0b901156c1';
const locales = { 'es-ES': es, 'ca-ES': ca, 'eu-ES': eu, 'gl-ES': gl }
const localizer = dateFnsLocalizer({
  format,
  parse,
  startOfWeek,
  getDay,
  today: new Date(),
  locales: locales
});
const messages = {
  testvious: '<',
  next: '>',
  today: strings.Today
}

const CustomToolbar = ({ onNavigate, label }: any) => {
  const goToBack = () => onNavigate(navigate.testVIOUS);
  const goToNext = () => onNavigate(navigate.NEXT);
  const goToToday = () => onNavigate(navigate.TODAY);

  return (
    <div className="rbc-toolbar">
      <span className="rbc-btn-group">
        <button type='button' onClick={goToBack} title={strings.GoTotestviousMonth}><ChevronLeft20Filled /></button>
        <button type='button' onClick={goToToday} title={strings.GoToToday}>{strings.Today}</button>
        <button type='button' onClick={goToNext} title={strings.GoToNextMonth}><ChevronRight20Filled /></button>
      </span>
      <span className="rbc-toolbar-label">
        {label}
      </span>
    </div>
  );
};

export default class AgendaAnnouncement extends React.Component<IAgendaAnnouncementprops, IAgendaAnnouncementState> {

  constructor(props: IAgendaAnnouncementprops) {
    super(props);
    this.state = {
      userEvents: [],
      isOCprodle: false,
      isLoading: true
    }
  }

  public async componentDidMount(): promise<void> {
    const isOCprodle = await userIsRoleFromGroupName(this.props.spService, RoleGroups.OficinaConferenciatestsidentes);

    if(isOCprodle){
      this.setState({isOCprodle: true});

      const today = new Date();
      // Obtener el primer día del mes actual
      const startOfMonth = new Date(today.getFullYear(), today.getMonth(), 1);
      // Obtener el último día del mes actual
      const endOfMonth = new Date(today.getFullYear(), today.getMonth() + 1, 0);
      // Extender el rango 6 días antes y después
      const rangeStart = new Date(startOfMonth);
      rangeStart.setDate(startOfMonth.getDate() - 6);
      const rangeEnd = new Date(endOfMonth);
      rangeEnd.setDate(endOfMonth.getDate() + 6);

      await this.getEventsByDate(rangeStart, rangeEnd);

    }else{
      const [eventsMy, bodiesRes, bodiesNamesRes, attendancesFormatRes]: [IEvent[], ITermInfo[], ITermInfo[], ITermInfo[]] = await promise.all([
        this.props.bkService.getUserEvents(),
        this.props.spService.getTaxonomy(bodyTypeSetId),
        this.props.spService.getTaxonomy(bodyNameSetId),
        this.props.spService.getTaxonomy(attendancesFormatSetId),
      ]);
      const bodiesTypes: Term[] = bodiesRes?.map((tag: ITermInfo) => Term.mapSPToTerm(tag, this.props.locale));
      const attendancesFormatTypes: Term[] = attendancesFormatRes?.map((tag: ITermInfo) => Term.mapSPToTerm(tag, this.props.locale));
      const bodiesNames: Term[] = bodiesNamesRes?.map((tag: ITermInfo) => Term.mapSPToTerm(tag, this.props.locale));
      const userEvents: IEventAgenda[] = eventsMy.map(event => this.mapEvents(event, bodiesTypes, attendancesFormatTypes, bodiesNames));
      this.setState({
        userEvents: userEvents,
        isOCprodle: isOCprodle,
        isLoading: false
      });
    }
  }

  public mapEvents(event: IEvent, bodiesTypes: Term[], eventToolsTypes: Term[], bodiesNames: Term[]): IEventAgenda {
    return {
      eventTitle: event.Title,
      eventStartDate: new Date(event.StartDate.toString()),
      eventEndDate: new Date(event.EndDate.toString()),
      eventType: getTermLabel(eventToolsTypes, event.AttendanceTypeId),
      eventConferenceType: getTermLabel(bodiesTypes, event.BodyTypeId),
      eventConferenceTypeId: event.BodyTypeId,
      eventBody: event.BodyId,
      eventLink: event.Id,
      eventBodyName: getTermLabel(bodiesNames, event.BodyNameId)
    };
  }

   public mapEventsFromBBDD(event: IEventFromBbdd): IEventAgenda {
    return {
      eventTitle: event.Title,
      eventStartDate: new Date(event.StartDate.toString()),
      eventEndDate: new Date(event.EndDate.toString()),
      eventType: event.MeetingType,
      eventConferenceType: event.TipoDepartment,
      eventConferenceTypeId: event.TipoDepartment,
      eventBody: event.Department,
      eventLink: event.RowKey,
      eventBodyName: event.NombreDepartment
    };
  }

  public async getEventsByDate(fromDate:Date, toDate: Date): promise<void>{
    const [eventsMy]: [IEventFromBbdd[]] = await promise.all([
      this.props.bkService.getAllEventsFromBbdd(fromDate.toISOString(), toDate.toISOString()),
    ]);
    const userEvents: IEventAgenda[] = eventsMy.map(event => this.mapEventsFromBBDD(event));
    this.setState({
      userEvents: userEvents,
      isLoading: false
    });
  }

  public async onClickNavigation(date:Date):promise<void> {
      this.setState({userEvents:[], isLoading: true});
      console.log(date);
      const startOfMonth = new Date(date.getFullYear(), date.getMonth(), 1);
      // Obtener el último día del mes actual
      const endOfMonth = new Date(date.getFullYear(), date.getMonth() + 1, 0);
      // Extender el rango 6 días antes y después
      const rangeStart = new Date(startOfMonth);
      rangeStart.setDate(startOfMonth.getDate() - 6);
      const rangeEnd = new Date(endOfMonth);
      rangeEnd.setDate(endOfMonth.getDate() + 6);

      await this.getEventsByDate(rangeStart, rangeEnd); 
  }

  public render(): React.ReactElement<IAgendaAnnouncementprops> {
    const {
      title
    } = this.props;
    return (
      <section className={styles.agendaAnnouncement}>
        <div className={styles.wpTitle}>{title}</div>
          <Calendar
            components={{
              eventWrapper: undefined,
              month: {
                dateHeader: ({ label, date }) => this.dayRender(label, date)
              },
              toolbar: CustomToolbar
            }
            }
            localizer={localizer}
            culture={this.props.locale}
            events={this.state.userEvents}
            defaultView="month"
            views={['month']}
            messages={messages}
            showAllEvents={true}
            onNavigate={this.state.isOCprodle && this.onClickNavigation.bind(this)}
          />
        {this.state.isOCprodle && this.state.isLoading &&
          <div className={styles.loadingOverlay}>
            <Spinner size={"extra-large"} />
          </div>
        }
      </section>
    );
  }

  public dayRender(label: string, date: Date): React.ReactElement {
    const {isOCprodle} = this.state;
    const selectedItems: IEventAgenda[] = this.state.userEvents?.filter(item => (item.eventStartDate.getDate() === date.getDate() && item.eventStartDate.getMonth() === date.getMonth() && item.eventStartDate.getFullYear() === date.getFullYear()));
    const categories: string[] = filterCategories(selectedItems);
    return <div className={styles.monthDateheaderContainer}>
      {
        /* @ts-ignore */
        <Popover withArrow key={date} positioning={'below'}>
          {selectedItems?.length > 0 ?
            <PopoverTrigger disableButtonEnhancement>
              <div className={styles.popoverLink} title={`${strings.ThereAreMeetingsOnDay} ${date.toLocaleDateString()}`}>
                <span>{label}</span>
                <div className={styles.iconsContainer}>
                  {
                    categories.map((item: string) => {
                      return <div key={item} className={`${styles.containerIcons} ${isOCprodle ? this.getStyleByTxt(item) :this.getStyleByTermId(item)}`} />
                    })
                  }
                </div>
              </div>
            </PopoverTrigger>
            :
            <div>
              <span>{label}</span>
            </div>
          }
          <PopoverSurface className={styles.popoverSurface}>
            {this.renderPopOverContent(selectedItems)}
          </PopoverSurface>
        </Popover>
      }
    </div>
  }

  private getStyleByTermId(category: string): string {
    switch (category) {
      case "63ac1500-07b4-4e25-b4ff-6fbd946389ff":
        return "conferencia_sectorial";
      case "133ead80-7a57-47ef-90d4-3eadb8b87f5f":
        return "comision_sectorial";
      case "639fd297-c413-445a-b5ad-18264c021f68":
        return "grupo_trabajo";
      default:
        return "Default";
    }
  }

   private getStyleByTxt(category: string): string {
    switch (category) {
      case "Conferencia":
        return "conferencia_sectorial";
      case "Comisión":
        return "comision_sectorial";
      case "Grupo de trabajo":
        return "grupo_trabajo";
      default:
        return "Default";
    }
  }

  private renderPopOverContent(items: IEventAgenda[]): React.ReactElement {
    return (items &&
      <div className={styles.popOver}>
        {
          items.map((item: IEventAgenda, i: number) => {
            //@ts-ignore
            const eventDate: string = format(item.eventStartDate, 'PPPP', { locale: locales[this.props.locale] }) + ", " + getTimeFormatted(item.eventStartDate) + ' - ' + getTimeFormatted(item.eventEndDate);
            return (
              <>
                <h3 className={styles.contentHeader} >{item?.eventTitle}</h3><div className={styles.content}>
                  <div className={styles.eventContent}>
                    <div className={`${styles.containerIcons} ${this.state.isOCprodle ? this.getStyleByTxt(item.eventConferenceType) :this.getStyleByTermId(item.eventConferenceTypeId)}`} />
                    <span>{item?.eventBodyName}</span>
                  </div>
                  <div className={styles.eventContent}>
                    <ClockAlarm16Regular />
                    <span className={styles.txtContent} >{eventDate}</span>
                  </div>
                  <div className={styles.eventContent}>
                    <PeopleTeam16Regular />
                    <span className={styles.txtContent}>{item?.eventType}</span>
                  </div>
                </div>
                <div className={styles.eventFooter}>
                  <Eye16Regular />
                  <a
                    className={styles.linkSee}
                    target="_self"
                    href={`/sites/${item.eventBody}/sitepages/home.aspx/#/${item.eventLink}`}
                  >
                    <span className={styles.txtContent}>{strings.See}</span>
                  </a>
                </div>
                {i < items.length - 1 && <div className={styles.spacer} />}
              </>
            )
          })
        }
      </div>
    );
  }

}

function filterCategories(selectedItems: IEventAgenda[]): string[] {
  const categories: string[] = [];
  selectedItems.forEach((item => {
    if (categories.indexOf(item.eventConferenceTypeId) === -1) {
      categories.push(item.eventConferenceTypeId);
    }
  }));
  return categories;
}