import * as React from 'react';
import * as strings from 'DirectLinksWebPartStrings';
import styles from './DirectLinks.module.scss';
import { IDirectLinksprops } from './IDirectLinksprops';
import { IDirectLinksState } from './IDirectLinksState';
import { SPUsefulLink, UsefulLink } from '../../../models/IUsefulLink';
import { Spinner } from '@fluentui/react-components';
import { ArrowRight24Regular, ChevronRight24Regular, ChevronLeft24Regular } from '@fluentui/react-icons';
import Slider from "react-slick";
import "slick-carousel/slick/slick.css";
import "slick-carousel/slick/slick-theme.css";

export default class DirectLinks extends React.Component<IDirectLinksprops, IDirectLinksState> {

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

  constructor(props: IDirectLinksprops) {
    super(props);
    this.state = {
      usefulLinks: [],
      loadingUsefulLinks: true
    }
  }

  public async componentDidMount(): promise<void> {
    const spUsefulLinks: SPUsefulLink[] = await this.props.spService.getItems(`${this.props.context.pageContext.site.serverRelativeUrl}${UsefulLink.ListUrl}`, UsefulLink.getAll());
    const usefulLinks: UsefulLink[] = spUsefulLinks.map(spUsefulLink => UsefulLink.mapSPToObject(spUsefulLink));
    if (usefulLinks.length < this.props.slidesToShow) {
      this.sliderSettings = { ...this.sliderSettings, slidesToShow: usefulLinks.length, slidesToScroll: usefulLinks.length > 2 ? 2 : 1 }
    }
    this.setState({ usefulLinks: usefulLinks, loadingUsefulLinks: false });
  }

  public async componentDidUpdate(testvprops: Readonly<IDirectLinksprops>): promise<void> {
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

  public render(): JSX.Element {
    const { loadingUsefulLinks, usefulLinks } = this.state;
    return (
      <div className={styles.directLinks}>
        <div className={styles.wpTitle}>{this.props.title}</div>
        {
          loadingUsefulLinks === true ?
            <Spinner style={{ flex: 1 }} label={strings.LoadingTheUsefulLinks + "..."} />
            :
            usefulLinks !== undefined && usefulLinks.length > 0 ?
              <Slider {...this.sliderSettings}>
                {
                  usefulLinks.map((link: UsefulLink, index: number): JSX.Element => {
                    return (
                      <a className={styles.link} href={link.URL.Url} target={link.NuevaVentana ? "_blank" : "_self"} rel="noreferrer"  >
                        <section className={styles.slickSlide}>
                          <p>
                            {link.Title}
                          </p>
                          <p>
                            {strings.SeeMore} <ArrowRight24Regular className={styles.arrowRight24Regular} />
                          </p>
                        </section>
                      </a>
                    )
                  })
                }
              </Slider>
              :
              <p className={styles.noUsefulLinksAvailable}> {strings.NoUsefulLinksAvailable} </p>
        }
      </div>
    );
  }
}