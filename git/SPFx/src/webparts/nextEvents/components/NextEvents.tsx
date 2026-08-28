/* eslint-disable @typescript-eslint/ban-ts-comment */
import * as React from 'react';
import * as strings from 'NextEventsWebPartStrings';
import styles from './NextEvents.module.scss';
import { INextEventsprops } from './INextEventsprops';
import { INextEventsState } from './INextEventsState';
import { Spinner } from '@fluentui/react-components';
import { ChevronRight24Regular, ChevronLeft24Regular } from '@fluentui/react-icons';
import Slider from "react-slick";
import "slick-carousel/slick/slick.css";
import "slick-carousel/slick/slick-theme.css";
import format from 'date-fns/format';
import es from 'date-fns/locale/es';
import ca from 'date-fns/locale/ca';
import eu from 'date-fns/locale/eu';
import gl from 'date-fns/locale/gl';
import { IEvent } from '../../../service/BackendServiceModels/EventModels';
import { ITermInfo } from '@pnp/sp/taxonomy';
import { getTimeFormatted } from '../../../utils/Utils';
import { Term } from '../../../models/ITag';

const bodyNameSetId: string = '6f98455a-a1c9-45e7-8c5d-aaca87690eff';
const attendancesFormatSetId: string = '0775b014-e8fc-4aac-ae49-6b0b901156c1';

export default class NextEvents extends React.Component<INextEventsprops, INextEventsState> {

  private sliderSettings = {
    nextArrow: <this.nextArrow right={true} />,
    testvArrow: <this.nextArrow right={false} />,
    dots: false,
    infinite: false,
    speed: 500,
    slidesToShow: this.props.slidesToShow,
    slidesToScroll: this.props.slidesToScroll,
    initialSlide: 0,
    responsive: [
      {
        breakpoint: 1024,
        settings: {
          dots: false,
          infinite: false,
          slidesToShow: this.props.slidesToShow,
          slidesToScroll: this.props.slidesToScroll,
        }
      },
      {
        breakpoint: 600,
        settings: {
          dots: false,
          infinite: false,
          slidesToShow: 2,
          slidesToScroll: 1
        }
      },
      {
        breakpoint: 480,
        settings: {
          dots: false,
          infinite: false,
          slidesToShow: 1,
          slidesToScroll: 1
        }
      }
    ]
  };

  constructor(props: INextEventsprops) {
    super(props);
    this.state = {
      nextEvents: [],
      loadingNextEvents: true,
      attendancesFormatTypes: [],
      bodiesNames: []
    }
  }

  public async componentDidMount(): promise<void> {
    const { bkService, spService, slidesToShow, locale } = this.props;
    let bodiesRes: ITermInfo[] = [];
    let bodiesNames: Term[] = [];
    let attendancesFormatTypes: Term[] = [];
    let attendancesFormatRes: ITermInfo[] = [];
    let myNextEvents: IEvent[] = [];
    try {
      [bodiesRes, attendancesFormatRes, myNextEvents] = await promise.all([
        spService.getTaxonomy(bodyNameSetId),
        spService.getTaxonomy(attendancesFormatSetId),
        bkService.getUserEvents()
      ]);
      bodiesNames = bodiesRes?.map((tag: ITermInfo) => Term.mapSPToTerm(tag, locale));
      attendancesFormatTypes = attendancesFormatRes?.map((tag: ITermInfo) => Term.mapSPToTerm(tag, locale));

      myNextEvents = myNextEvents.filter((event) => {
        return new Date(event.StartDate.toString()) > new Date()
      });
    }
    catch (err) {
      console.log(err);
    }

    if (myNextEvents.length < slidesToShow) {
      this.sliderSettings = { ...this.sliderSettings, slidesToShow: myNextEvents.length, slidesToScroll: myNextEvents.length > 2 ? 2 : 1 }
    }
    this.setState({ nextEvents: myNextEvents, loadingNextEvents: false, bodiesNames, attendancesFormatTypes });
  }

  public async componentDidUpdate(testvprops: Readonly<INextEventsprops>): promise<void> {
    if (testvprops.slidesToShow !== this.props.slidesToShow || testvprops.slidesToScroll !== this.props.slidesToScroll) {
      this.sliderSettings = { ...this.sliderSettings, slidesToShow: this.props.slidesToShow, slidesToScroll: this.props.slidesToScroll }
      this.setState(this.state);
    }
  }

  private nextArrow(props: { className?: string; onClick?: React.MouseEventHandler<HTMLDivElement>; right: boolean; }): JSX.Element {
    const { className, onClick, right } = props;
    return (
      <div title={right ? strings.MoveRight : strings.MoveLeft}
        className={`${className} ${styles.arrow}`}
        onClick={onClick}
      >
        {right ? <ChevronRight24Regular /> : <ChevronLeft24Regular />}
      </div>
    );
  }

  public renderMoreLink = (): React.ReactElement => {
    return <a className={styles.moreLink} href={this.props.moreLink} title={strings.ViewMoreUpcomingEvents}>{strings.MoreLink}</a>;
  };

  public render(): JSX.Element {
    const { loadingNextEvents, nextEvents, bodiesNames, attendancesFormatTypes } = this.state;
    const locales = { 'es-ES': es, 'ca-ES': ca, 'eu-ES': eu, 'gl-ES': gl };
    return (
      <div className={styles.nextEvents}>
        <div className={styles.titleRow}>
          <div className={styles.wpTitle}>{this.props.title}</div>
          {
            this.props.moreLink && this.renderMoreLink()
          }
        </div>
        {
          loadingNextEvents === true ?
            <Spinner style={{ flex: 1 }} label={strings.LoadingTheNextEvents + "..."} />
            :
            nextEvents !== undefined && nextEvents.length > 0 ?
              <Slider {...this.sliderSettings}>
                {
                  nextEvents.map((event: IEvent): JSX.Element => {
                    const startDate = new Date(event.StartDate);
                    const endDate = new Date(event.EndDate);
                    const attendanceType = attendancesFormatTypes.find(term => term.key === event.AttendanceTypeId);
                    const bodyName = bodiesNames.find(term => term.key === event.BodyNameId);

                    /* @ts-ignore */
                    const eventDateAux: string = format(startDate, 'PPPP', { locale: locales[this.props.locale] }) + ", " + getTimeFormatted(startDate) + ' - ' + getTimeFormatted(endDate);
                    const eventDate: string = eventDateAux.charAt(0).toUpperCase() + eventDateAux.slice(1);
                    return (
                      <a className={styles.link}>
                        <section className={styles.slickSlide}>
                          <section className={styles.slickSlide_firstContainer}>
                            <p>
                              {
                                /* @ts-ignore */
                                format(startDate, 'MMM', { locale: locales[this.props.locale] }).toUpperCase()
                              }
                            </p>
                            <p> {startDate.getDate()} </p>
                          </section>
                          <section className={styles.slickSlide_secondContainer}>
                            <p> {bodyName?.text} </p>
                            <p> {event.Title} </p>
                            <p> {eventDate} </p>
                            <p> {attendanceType?.text} </p>
                          </section>
                          <section className={styles.slickSlide_thirdContainer}>
                            <p className={styles.GoToEvent}>
                              <a
                                title={`${strings.GoToTheEventDetail}: ${event.Title}`}
                                target="_self"
                                href={`/sites/${event.BodyId}/sitepages/home.aspx/#/${event.Id}`}
                              >
                                {strings.GoToEvent}
                              </a>
                            </p>
                          </section>
                        </section>
                      </a>
                    )
                  })
                }
              </Slider>
              :
              <p className={styles.noNextEventsAvailable}> {strings.NoNextEventsAvailable} </p>
        }
      </div>
    );
  }

}