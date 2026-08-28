import * as React from 'react';
import * as strings from 'myDepartmentsWebPartStrings';
import styles from './myDepartments.module.scss';
import { ImyDepartmentsprops } from './ImyDepartmentsprops';
import { ImyDepartmentsState } from './ImyDepartmentsState';
import { Spinner } from '@fluentui/react-components';
import { ChevronRight24Regular, ChevronLeft24Regular } from '@fluentui/react-icons';
import Slider from "react-slick";
import "slick-carousel/slick/slick.css";
import "slick-carousel/slick/slick-theme.css";
import { IBody } from '../../../service/BackendServiceModels/BodyModel';
import { ITermInfo } from '@pnp/sp/taxonomy';
import { Term } from '../../../models/ITag';
import { getTermLabel } from '../../../utils/Utils';

const bodyTypeSetId: string = 'ba3ff2ef-09c8-4345-9bad-777014ddf636';

export default class myDepartments extends React.Component<ImyDepartmentsprops, ImyDepartmentsState> {

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

  constructor(props: ImyDepartmentsprops) {
    super(props);
    this.state = {
      myDepartments: [],
      loadingmyDepartments: true,
      bodiesTypes: []
    }
  }

  public async componentDidMount(): promise<void> {
    const { bkService, spService, slidesToShow, locale } = this.props;
    try {
      const [myDepartments, bodiesRes]: [IBody[], ITermInfo[]] = await promise.all([
        bkService.getUserBodies(),
        spService.getTaxonomy(bodyTypeSetId)
      ]);
      const bodiesTypes: Term[] = bodiesRes?.map((tag: ITermInfo) => Term.mapSPToTerm(tag, locale));
      if (myDepartments.length < slidesToShow) {
        this.sliderSettings = { ...this.sliderSettings, slidesToShow: myDepartments.length, slidesToScroll: (myDepartments.length > 2) ? 2 : 1 }
      }
      this.setState({ myDepartments: myDepartments, loadingmyDepartments: false, bodiesTypes });
    }
    catch (error) {
      console.log(error);
      this.setState({ loadingmyDepartments: false });
    }
  }

  public async componentDidUpdate(testvprops: Readonly<ImyDepartmentsprops>): promise<void> {
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

  private getColorForOrganType(organType: string): string {
    switch (organType) {
      case "133ead80-7a57-47ef-90d4-3eadb8b87f5f":   //comision
        return "#B58526";
      case "63ac1500-07b4-4e25-b4ff-6fbd946389ff":   //conferencia
        return "#39579D";
      case "639fd297-c413-445a-b5ad-18264c021f68":   //grupo trabajo
        return "#858099";
      default:
        return "gray";
    }
  }

  private getOrganAbbreviaton(organType: string): string {
    switch (organType) {
      case "133ead80-7a57-47ef-90d4-3eadb8b87f5f":   //comision
        return "CO";
      case "63ac1500-07b4-4e25-b4ff-6fbd946389ff":   //conferencia
        return "CS";
      case "639fd297-c413-445a-b5ad-18264c021f68":   //grupo trabajo
        return "GR";
      default:
        return "";
    }
  }

  public render(): JSX.Element {
    const { loadingmyDepartments, myDepartments } = this.state;
    return (
      <div className={styles.myDepartments}>
        <div className={styles.wpTitle}>{this.props.title}</div>
        {
          loadingmyDepartments === true ?
            <Spinner style={{ flex: 1 }} label={strings.LoadingmyDepartments + "..."} />
            :
            myDepartments !== undefined && this.state.myDepartments.length > 0 ?
              <Slider {...this.sliderSettings}>
                {
                  this.state.myDepartments.map((organ: IBody, index: number): JSX.Element => {
                    return (
                      <a
                        className={styles.link}
                        key={index}
                        target="_self"
                        href={`${organ.RelativeUrl}/sitepages/home.aspx/`}
                      >
                        <section className={styles.slickSlide} key={index} >
                          <section className={styles.slickSlide_firstContainer} style={{ backgroundColor: this.getColorForOrganType(organ.TypeId) }}>
                            <section className={styles.slickSlide_firstSubContainer}>
                              <section>
                                <section className={styles.siteLogoUrl} style={{ backgroundColor: this.getColorForOrganType(organ.TypeId) }}>
                                  {this.getOrganAbbreviaton(organ.TypeId)}
                                </section>
                              </section>
                              <section>
                                <p> {organ.Title} </p>
                              </section>
                              <section>
                                <p> {getTermLabel(this.state.bodiesTypes, organ.TypeId)} </p>
                              </section>
                            </section>
                          </section>
                          <section className={styles.slickSlide_secondContainer}>
                            <p> {organ.Description} </p>
                          </section>
                        </section>
                      </a>
                    )
                  })
                }
              </Slider>
              :
              <p className={styles.nomyDepartmentsAvailable}> {strings.NomyDepartmentsAvailable} </p>
        }
      </div>
    );
  }


}