/* eslint-disable @typescript-eslint/ban-ts-comment */
/* eslint-disable no-void */
import * as React from 'react';
import * as strings from 'WelcomeMessageWebPartStrings';
import * as yup from 'yup';
import styles from './WelcomeMessage.module.scss';
import { IWelcomeMessageprops, IWelcomeMessageState } from './IWelcomeMessage';
import {
  Button,
  Dialog,
  DialogActions,
  DialogBody,
  DialogContent,
  DialogSurface,
  DialogTitle,
  Dropdown,
  Input,
  SkeletonItem,
  Option,
  Popover,
  PopoverSurface,
  PopoverTrigger,
  Link,
  Skeleton,
  Listbox,
  MessageBar,
  MessageBarBody,
  MessageBarTitle,
  Field,
  Spinner,
  Text,
  Accordion,
  AccordionItem,
  AccordionHeader,
  AccordionPanel,
  Divider,
  Drawer,
  DrawerBody,
  DrawerHeader,
  DrawerHeaderTitle,
  Tooltip
} from '@fluentui/react-components';
import {
  AlertOn20Regular,
  Briefcase20Regular,
  CircleSmall20Regular,
  Dismiss24Regular,
  Edit24Regular,
  Info24Regular,
  Mail20Regular,
  MailSettings20Regular,
  NavigationLocationTarget20Regular,
  Person20Regular,
  Phone20Regular,
  TextBulletListSquareWarning20Regular
} from '@fluentui/react-icons';
import { TipoAlert } from '../../../models/IAlert';
import { ITermInfo } from '@pnp/sp/taxonomy';
import { Term } from '../../../models/ITag';
import { INotification, NotificationBatchOp } from '../../../service/BackendServiceModels/NotificationModel';
import { Iprofile } from '../../../service/BackendServiceModels/profileModel';
import { Form, Formik } from 'formik';
import { getTermLabel, getTime } from '../../../utils/Utils';
import { IPendingChanges } from '../../../models/IPendingChanges';
import { Person } from '@microsoft/mgt-react/dist/es6/spfx';
import { ViewType } from '@microsoft/mgt-spfx';
import format from 'date-fns/format';
import es from 'date-fns/locale/es';
import ca from 'date-fns/locale/ca';
import eu from 'date-fns/locale/eu';
import gl from 'date-fns/locale/gl';

const AutonomousCommunitiesSetId: string = '172d84e2-ba88-40d6-a872-8fe00d066caa';
const AtendanceTypeSetId: string = '0775b014-e8fc-4aac-ae49-6b0b901156c1';
const OnlineToolSetId: string = '55dfa3e7-f0df-4dc8-b53c-4018418041fe';

export default class WelcomeMessage extends React.Component<IWelcomeMessageprops, IWelcomeMessageState> {

  currentAlertsReferences: any[];
  observers: IntersectionObserver[];

  constructor(props: IWelcomeMessageprops) {
    super(props);
    this.state = {
      loadingUserInformation: false,
      savingUserInformation: false,
      loadingArchivedAlerts: false,
      autonomousCommunities: [],
      attendanceTypes: [],
      onlineTools: [],
      dialogIsOpen: false,
      popoverIsOpen: false,
      drawerIsOpen: false,
      backendErrorMessage: "",
      backendHasFailed: false,
      pendingChanges: undefined,
      selectedCurrentAlert: undefined,
      selectedArchivedAlert: undefined,
      archivingMeetings: false
    };
    this.currentAlertsReferences = [];
  }

  public componentDidMount(): void {
    void this.onInit();
  }

  private async onInit(): promise<void> {
    const { context, spService, bkService, getNotificationsFromAllSites } = this.props;
    const source = context.pageContext.web.serverRelativeUrl.split("sites/")[1];
    const locale = context.pageContext.cultureInfo.currentUICultureName;
    try {
      /* Se obtienen los Alerts actuales, los Alerts archivados y las taxonomías necesarias */
      const [currentAlerts, autonomousCommunitiesRes, attendanceTypesRes, onlineToolsRes]: [INotification[], ITermInfo[], ITermInfo[], ITermInfo[]] = await promise.all([
        bkService.getUserNotifications(getNotificationsFromAllSites ? "" : source, locale),
        spService.getTaxonomy(AutonomousCommunitiesSetId),
        spService.getTaxonomy(AtendanceTypeSetId),
        spService.getTaxonomy(OnlineToolSetId)
      ]);
      const sortedCurrentAlerts: INotification[] = [...currentAlerts].sort((a: INotification, b: INotification): number => {
        return getTime(b.Created) - getTime(a.Created);
      });
      // Se filtran los Alerts que son más antiguos que la fecha actual
      const currentAlertsToShow: INotification[] = await this.hideExpiredAlerts(sortedCurrentAlerts);
      currentAlertsToShow.forEach((): void => {
        this.currentAlertsReferences.push(React.createRef());
      });
      const autonomousCommunities: Term[] = autonomousCommunitiesRes.map((tag: ITermInfo): Term =>
        Term.mapSPToTerm(tag, context.pageContext.cultureInfo.currentUICultureName)
      );
      const attendanceTypes: Term[] = attendanceTypesRes.map((tag: ITermInfo): Term =>
        Term.mapSPToTerm(tag, context.pageContext.cultureInfo.currentUICultureName)
      );
      const onlineTools: Term[] = onlineToolsRes.map((tag: ITermInfo): Term =>
        Term.mapSPToTerm(tag, context.pageContext.cultureInfo.currentUICultureName)
      );
      this.setState({
        currentAlerts: currentAlertsToShow,
        autonomousCommunities,
        attendanceTypes,
        onlineTools
      });
    }
    catch (error) {
      console.error(error);
    }
  }

  private registerObserver(): void {
    const { currentAlerts } = this.state;
    if (currentAlerts) {
      this.observers = currentAlerts.map((currentAlert: INotification, index: number): IntersectionObserver => {
        const observer: IntersectionObserver = new IntersectionObserver(async ([entry]: IntersectionObserverEntry[]): promise<void> => {
          if (entry.isIntersecting) {
            await this.setCurrentAlertAsRead(entry.target.id);
          }
        });
        const currentAlertReference = this.currentAlertsReferences[index];
        if (currentAlertReference && currentAlertReference.current) {
          observer.observe(currentAlertReference.current as HTMLDivElement);
        }
        return observer;
      });
    }
  }

  private async setCurrentAlertAsRead(currentAlertId: string): promise<void> {
    const { bkService } = this.props;
    const { currentAlerts } = this.state;
    if (currentAlerts) {
      const currentAlertIndex: number = currentAlerts.findIndex((currentAlert: INotification): boolean => currentAlert.Id === currentAlertId);
      if ((currentAlertIndex > -1) && (currentAlerts[currentAlertIndex].Visualized === false)) {
        try {
          await bkService.setReadNotification(currentAlertId, true);
          const newCurrentAlerts = currentAlerts?.map((currentAlert: INotification, index: number): INotification => {
            if (index === currentAlertIndex) {
              return { ...currentAlert, Visualized: true };
            }
            return currentAlert;
          })
          this.setState({ currentAlerts: newCurrentAlerts });
        }
        catch (error) {
          console.error(error);
        }
      }
    }
  }

  private async hideExpiredAlerts(allAlerts: INotification[]): promise<INotification[]> {
    const { bkService } = this.props;
    const expiredAlerts: INotification[] = [];
    const currentAlertsToShow: INotification[] = allAlerts.filter((alert: INotification): boolean => {
      let hasExpired: boolean = false;
      if (alert.Expires) {
        alert.Expires = new Date(alert.Expires);
        if (
          (new Date() > alert.Expires) &&
          (alert.Expires.getTime() !== new Date("0001-01-01T00:00:00").getTime())
        ) {
          expiredAlerts.push(alert);
          hasExpired = true;
        }
      }
      return !hasExpired;
    });
    if (expiredAlerts.length > 0) {
      try {
        await bkService.updateNotificationsInBatch(
          {
            Operation: NotificationBatchOp.Hide,
            NotificationIds: expiredAlerts.map((expiredAlert: INotification): string => {
              return expiredAlert.Id;
            })
          }
        );
      }
      catch (error) {
        console.error(error);
      }
    }
    return currentAlertsToShow || [];
  }

  private async archiveCurrentReadAlerts(): promise<void> {
    const { bkService } = this.props;
    const { currentAlerts } = this.state;
    const currentAlertsToArchive: INotification[] | undefined = currentAlerts?.filter((currentAlert: INotification): boolean => currentAlert.Visualized);
    const newCurrentAlerts: INotification[] | undefined = currentAlerts?.filter((currentAlert: INotification): boolean => !currentAlert.Visualized);
    this.setState({ archivingMeetings: true });
    if (currentAlertsToArchive) {
      try {
        await bkService.updateNotificationsInBatch({
          Operation: NotificationBatchOp.Hide,
          NotificationIds: currentAlertsToArchive.map((currentAlert: INotification): string => {
            return currentAlert.Id
          })
        });
      }
      catch (error) {
        console.error(error);
      }
      finally {
        this.setState({ archivingMeetings: false });
      }
    }
    this.setState({ currentAlerts: newCurrentAlerts });
  }

  private async getUserInformation(): promise<void> {
    const { bkService } = this.props;
    this.setState({
      dialogIsOpen: true,
      loadingUserInformation: true,
      savingUserInformation: false
    });
    try {
      const userInformation: Iprofile = await bkService.getCurrentUserprofile();
      this.setState({
        userInformation,
        loadingUserInformation: false
      });
    }
    catch (error) {
      console.error(error);
      this.setState({ loadingUserInformation: false });
    }
  }

  private openDialog = async (): promise<void> => {
    await this.getUserInformation();
  };

  private closeDialog = (): void => {
    this.setState({
      dialogIsOpen: false,
      backendErrorMessage: "",
      backendHasFailed: false
    });
  };

  private saveUserInformation = async (userInformation: Iprofile): promise<void> => {
    const { bkService } = this.props;
    if (userInformation) {
      this.setState({ savingUserInformation: true });
      try {
        await bkService.updateCurrentUserprofile(userInformation);
        this.setState({
          savingUserInformation: false,
          dialogIsOpen: false
        });
      }
      catch (error) {
        console.error(error);
        this.setState({
          savingUserInformation: false,
          backendErrorMessage: strings.BackendErrorMessage,
          backendHasFailed: true
        });
      }
    }
  };

  private onPopoverChange(newPopoverStatus: boolean): void {
    if (newPopoverStatus) {
      this.setState({ popoverIsOpen: true }, (): void => this.registerObserver());
    }
    else {
      this.observers.forEach((observer: IntersectionObserver): void => {
        if (observer) {
          observer.disconnect();
        }
      });
      this.setState({ popoverIsOpen: false });
    }
  }

  private onRenderSkeleton(): React.ReactElement {
    return (
      <Skeleton className={styles.content}>
        {
          [...Array(6)].map((index: number): JSX.Element => (
            <SkeletonItem key={index} shape="rectangle" size={32} />
          ))
        }
      </Skeleton>
    );
  }

  private onRenderForm(): React.ReactElement {
    const { userInformation, savingUserInformation, autonomousCommunities, backendErrorMessage } = this.state;
    const initialFormFieldValues: Iprofile = userInformation || {
      FirstName: "",
      LastName: "",
      Email: "",
      CellPhone: "",
      JobTitle: "",
      Office: "",
      PrincipalMail: "",
      BusinessPhone: ""
    }
    const validationSchema = yup.object().shape({
      Email: yup
        .string()
        .matches(/^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/, strings.AValidEmailIsRequired),
      CellPhone: yup
        .string()
        .matches(/^\d{9}$/, strings.AValidPhoneNumberIsRequired)
    });
    return (
      <div className={styles.content}>
        <MessageBar intent={"warning"} style={{ padding: "10px", marginBottom: "5px" }}>
          <MessageBarBody>
            <MessageBarTitle> {strings.Warning} </MessageBarTitle>
            {strings.WarningMessage}
          </MessageBarBody>
        </MessageBar>
        <Formik
          enableReinitialize
          initialValues={initialFormFieldValues}
          validationSchema={validationSchema}
          validateOnChange={false}
          validateOnBlur={false}
          onSubmit={(values: Iprofile): promise<void> => this.saveUserInformation(values)}
        >
          {
            (formik: any): JSX.Element => {
              const {
                errors,
                handleSubmit,
                values,
                handleChange,
                setFieldValue
              } = formik;
              return (
                <Form id="details" className={styles.formContainer} noValidate onSubmit={handleSubmit}>
                  <div className={styles.contentRow}>
                    <Tooltip content={strings.FirstNameTooltip} relationship='label' withArrow positioning={'below'}>
                      <Person20Regular />
                    </Tooltip>
                    <Input
                      id={"FirstName"}
                      className={styles.contentRowImput}
                      disabled
                      value={values.FirstName}
                      onChange={handleChange}
                    />
                  </div>
                  <div className={styles.contentRow}>
                    <Tooltip content={strings.LastNameTooltip} relationship='label' withArrow positioning={'below'}>
                      <Person20Regular />
                    </Tooltip>
                    <Input
                      id={"LastName"}
                      className={styles.contentRowImput}
                      disabled
                      value={values.LastName}
                      onChange={handleChange}
                    />
                  </div>
                  <div className={styles.contentRow}>
                    <Tooltip content={strings.PrincipalMailTooltip} relationship='label' withArrow positioning={'below'}>
                      <MailSettings20Regular />
                    </Tooltip>
                    <Field validationMessage={errors.Email} className={styles.contentRowImput}>
                      <Input
                        id={"PrincipalMail"}
                        className={styles.contentRowImput}
                        disabled={true}
                        value={values.PrincipalMail}
                        placeholder={strings.Email}
                        type="email"
                        onChange={handleChange}
                      />
                    </Field>
                  </div>
                  <div className={styles.contentRow}>
                    <Tooltip content={strings.PhoneTooltip} relationship='label' withArrow positioning={'below'}>
                      <Phone20Regular />
                    </Tooltip>
                    <Field validationMessage={errors.CellPhone} className={styles.contentRowImput}>
                      <Input
                        id={"CellPhone"}
                        className={styles.contentRowImput}
                        disabled={savingUserInformation}
                        value={values.CellPhone}
                        placeholder={strings.Phone}
                        type="tel"
                        onChange={handleChange}
                      />
                    </Field>
                  </div>
                  <div className={styles.contentRow}>
                    <Tooltip content={strings.OtherMailTooltip} relationship='label' withArrow positioning={'below'}>
                      <Mail20Regular />
                    </Tooltip>
                    <Field validationMessage={errors.Email} className={styles.contentRowImput}>
                      <Input
                        id={"Email"}
                        className={styles.contentRowImput}
                        disabled={savingUserInformation}
                        value={values.Email}
                        placeholder={strings.Email}
                        type="email"
                        onChange={handleChange}
                      />
                    </Field>
                  </div>
                  <div className={styles.contentRow}>
                    <Tooltip content={strings.JobTitleTooltip} relationship='label' withArrow positioning={'below'}>
                      <Briefcase20Regular />
                    </Tooltip>
                    <Input
                      id={"JobTitle"}
                      className={styles.contentRowImput}
                      disabled={true}
                      value={values.JobTitle}
                      placeholder={strings.Job}
                      onChange={handleChange}
                    />
                  </div>
                  <div className={styles.contentRow}>
                    <Tooltip content={strings.OfficeLocationTooltip} relationship='label' withArrow positioning={'below'}>
                      <NavigationLocationTarget20Regular />
                    </Tooltip>
                    <Dropdown
                      id={"Office"}
                      className={styles.contentRowImput}
                      disabled={true}
                      value={values.Office}
                      placeholder={strings.Location}
                      selectedOptions={values ? [values.Office] : []}
                      onOptionSelect={(_ev, data): void => setFieldValue("Office", data.optionText)}
                    >
                      <Listbox style={{ maxHeight: "200px" }}>
                        {autonomousCommunities.map((autonomousCommunity: Term): JSX.Element =>
                          <Option key={autonomousCommunity.key} value={autonomousCommunity.key}>{autonomousCommunity.text}</Option>)}
                      </Listbox>
                    </Dropdown>
                  </div>
                  {
                    (backendErrorMessage && backendErrorMessage !== "") &&
                    <div>
                      <span style={{ color: "red" }}> {backendErrorMessage} </span>
                    </div>
                  }
                </Form>
              );
            }}
        </Formik>
      </div>
    );
  }

  private onClickAlert(alert: INotification): void {
    if (alert && alert.Url) {
      const splittedAlertUrl: string[] = alert.Url.split("#");
      const modifiedAlertUrl: string = `${splittedAlertUrl[0]}/sitepages/home.aspx#${splittedAlertUrl[1]}`;
      history.pushState({}, "", modifiedAlertUrl);
      window.location.reload();
    }
  }

  private getTheNumberOfReadAndUnreadAlerts(alertsRead: boolean): number {
    const { currentAlerts } = this.state;
    return (
      currentAlerts?.filter((currentAlert: INotification): boolean =>
        currentAlert.Visualized === alertsRead)?.length ?? 0
    );
  }

  private filterPendingChanges(changeType: string, pendingChanges?: IPendingChanges[]): IPendingChanges[] | undefined {
    return pendingChanges?.filter((pendingChange: IPendingChanges): boolean => pendingChange.ChangeType === changeType);
  }

  private async handleDrawerOpening(): promise<void> {
    const { context, bkService, getNotificationsFromAllSites } = this.props;
    const { archivedAlerts } = this.state;
    const source = context.pageContext.web.serverRelativeUrl.split("sites/")[1];
    const locale = context.pageContext.cultureInfo.currentUICultureName;
    this.setState({ drawerIsOpen: true });
    if (!(archivedAlerts && archivedAlerts.length > 0)) {
      try {
        this.setState({ loadingArchivedAlerts: true });
        const archivedAlerts: INotification[] = await bkService.getArchivedNotifications(getNotificationsFromAllSites ? "" : source, locale);
        const sortedArchivedAlerts: INotification[] = [...archivedAlerts].sort((a: INotification, b: INotification): number => {
          return getTime(b.Created) - getTime(a.Created);
        });
        this.setState({ archivedAlerts: sortedArchivedAlerts, loadingArchivedAlerts: false });
      }
      catch (error) {
        console.error(error);
      }
    }
  }

  private renderAlerts(): React.ReactElement | undefined {
    const { locale } = this.props;
    const { currentAlerts, popoverIsOpen, attendanceTypes, onlineTools, selectedCurrentAlert, archivingMeetings } = this.state;
    const locales = { "es-ES": es, "ca-ES": ca, "eu-ES": eu, "gl-ES": gl };
    if (currentAlerts) {
      return (
        <div className={styles.multipleAlerts}>
          {/* Mensaje para indicar los Alerts pendientes */}
          <Info24Regular className={styles.alertIcon} />
          <span className={styles.alertBody}>
            {
              (this.getTheNumberOfReadAndUnreadAlerts(false) === 1) ?
                `${strings.AlertPhrase1} ${this.getTheNumberOfReadAndUnreadAlerts(false)} ${strings.AlertPhrase2}, `
                :
                `${strings.AlertPhrase1} ${this.getTheNumberOfReadAndUnreadAlerts(false)} ${strings.AlertPhrase3}, `
            }
            {/* @ts-ignore */}
            <Popover
              withArrow
              positioning={'below-start'}
              open={popoverIsOpen}
              onOpenChange={(e, data): void => this.onPopoverChange(data.open)}
            >
              <PopoverTrigger>
                <span>
                  <Link className={styles.alertLink}>
                    {strings.AlertPhrase4}
                  </Link>.
                </span>
              </PopoverTrigger>
              <PopoverSurface className={styles.popoverContainer}>
                <React.Fragment>
                  {/* Título del popover */}
                  <h3 style={{ marginTop: "0", padding: "0 12px" }}>
                    {strings.Alerts + " (" + currentAlerts.length + ")"}
                  </h3>
                  {
                    /* Botones del popover */
                    <div className={styles.popoverButtons}>
                      {/* Botón para ver los Alerts archivados */}
                      <Button
                        appearance="subtle"
                        onClick={(): promise<void> => this.handleDrawerOpening()}
                      >
                        {strings.ViewArchivedAlerts}
                      </Button>
                      {/* Botón para archivar los Alerts leidos */}
                      <div className={styles.archiveButton}>
                        {
                          archivingMeetings &&
                          <Spinner size='tiny' />
                        }
                        <Button
                          appearance="subtle"
                          disabled={this.getTheNumberOfReadAndUnreadAlerts(true) === 0 || archivingMeetings}
                          onClick={(): promise<void> => this.archiveCurrentReadAlerts()}
                        >
                          {strings.ArchiveReadAlerts}
                        </Button>
                      </div>
                    </div>
                  }
                  {
                    /* Cajón lateral con los Alerts archivados */
                    this.renderTheArchivedAlertsDrawer()
                  }
                </React.Fragment>
                {
                  (currentAlerts.length > 0) ?
                    <div>
                      {
                        /* Se muestran los Alerts */
                        currentAlerts.map((currentAlert: INotification, index: number): JSX.Element => {
                          const currentAlertReference = this.currentAlertsReferences[index];

                          /* Filtrado de cambios en los datos básicos de la Meeting */
                          const title: string | undefined = this.filterPendingChanges("Title", currentAlert.PendingChanges)?.[0]?.ChangeValue;
                          const location: string | undefined = this.filterPendingChanges("Location", currentAlert.PendingChanges)?.[0]?.ChangeValue;
                          const locationInformation: string | undefined = this.filterPendingChanges("LocationDetails", currentAlert.PendingChanges)?.[0]?.ChangeValue;
                          const meetingType: string | undefined = this.filterPendingChanges("MeetingType", currentAlert.PendingChanges)?.[0]?.ChangeValue;
                          const description: string | undefined = this.filterPendingChanges("Description", currentAlert.PendingChanges)?.[0]?.ChangeValue;
                          const startDate: string | undefined = this.filterPendingChanges("StartDate", currentAlert.PendingChanges)?.[0]?.ChangeValue;
                          const endDate: string | undefined = this.filterPendingChanges("EndDate", currentAlert.PendingChanges)?.[0]?.ChangeValue;
                          const onlineTool: string | undefined = this.filterPendingChanges("OnlineTool", currentAlert.PendingChanges)?.[0]?.ChangeValue;
                          const onlineToolUrl: string | undefined = this.filterPendingChanges("UrlOnlineTool", currentAlert.PendingChanges)?.[0]?.ChangeValue;

                          /* Filtrado de cambios en el orden del día */
                          const agendaPoints = this.filterPendingChanges("AgendaItem", currentAlert.PendingChanges);

                          /* Filtrado de cambios en los asistentes */
                          const addedAssistants: IPendingChanges[] | undefined = this.filterPendingChanges("AddAttendance", currentAlert.PendingChanges);
                          const removedAssistants: IPendingChanges[] | undefined = this.filterPendingChanges("RemoveAttendance", currentAlert.PendingChanges);

                          /* Filtrado de cambios en los Documents */
                          const documents: IPendingChanges[] | undefined = this.filterPendingChanges("Document", currentAlert.PendingChanges);

                          /* Filtrado de cambios en los acuerdos */
                          const agreements: IPendingChanges[] | undefined = this.filterPendingChanges("Agreement", currentAlert.PendingChanges);

                          /* Filtrado de cambios en el Minutes */
                          const minutes: IPendingChanges[] | undefined = this.filterPendingChanges("Minutes", currentAlert.PendingChanges);

                          return (
                            <>
                              <div
                                className={styles.popoverRow}
                                id={currentAlert.Id}
                                tabIndex={-1}
                                key={index}
                                ref={currentAlertReference}
                              >
                                <div className={styles.popoverTitle}>
                                  {
                                    (currentAlert.NotificationType === TipoAlert.Sistema) ?
                                      <TextBulletListSquareWarning20Regular
                                        className={styles.alertIcon}
                                        style={{ minHeight: "20px", minWidth: "20px" }}
                                      />
                                      :
                                      <AlertOn20Regular
                                        className={styles.alertIcon}
                                        style={{ minHeight: "20px", minWidth: "20px" }}
                                      />
                                  }
                                  <span
                                    className={currentAlert.Visualized ? `${styles.alertTitle} ${styles.readed}` : `${styles.alertTitle} ${styles.notReaded}`}
                                    style={{ cursor: "pointer", width: "100%" }}
                                    onClick={(): void => this.onClickAlert(currentAlert)}
                                    title={strings.GoToTheMeeting + " " + currentAlert.Title}
                                  >
                                    {currentAlert.Title}
                                  </span>
                                </div>
                                <span className={styles.alertBody} style={{ display: "flex", flexDirection: "column" }}>
                                  <span
                                    style={{ cursor: "pointer" }}
                                    onClick={(): void => this.onClickAlert(currentAlert)}
                                    title={strings.GoToTheMeeting + " " + currentAlert.Title}
                                  >
                                    {currentAlert.Body}
                                  </span>
                                  {
                                    (currentAlert.PendingChanges && currentAlert.PendingChanges.length > 0) ?
                                      (selectedCurrentAlert === currentAlert) ?
                                        <span>
                                          <Link
                                            style={{ width: "fit-content" }}
                                            onClick={(): void => this.setState({ selectedCurrentAlert: undefined })}
                                          >
                                            {strings.SeeLess}
                                          </Link>.
                                        </span>
                                        :
                                        <span>
                                          <Link
                                            style={{ width: "fit-content" }}
                                            onClick={(): void => this.setState({ selectedCurrentAlert: currentAlert })}
                                          >
                                            {strings.SeeMore}
                                          </Link>.
                                        </span>
                                      :
                                      undefined
                                  }
                                </span>
                                {
                                  /* Desplegable de un Alert asociado a una Meeting actualizada */
                                  selectedCurrentAlert === currentAlert &&
                                  <Accordion collapsible style={{ display: "flex", flexDirection: "column", paddingLeft: "13px" }}>
                                    {
                                      /* Cambios en los datos básicos de la Meeting */
                                      (title || location || locationInformation || meetingType || description || startDate || endDate || onlineTool || onlineToolUrl) && (
                                        <AccordionItem value="1">
                                          <AccordionHeader>
                                            <Field>
                                              <span>
                                                {strings.BasicData + ":"}
                                              </span>
                                            </Field>
                                          </AccordionHeader>
                                          <AccordionPanel style={{ display: "flex", flexDirection: "column" }}>
                                            <Field style={{ paddingLeft: "27px" }}>
                                              {
                                                title &&
                                                <Field style={{ display: "flex", alignItems: "flex-start" }}>
                                                  <CircleSmall20Regular style={{ minWidth: "20px" }} />
                                                  <span>
                                                    <strong>
                                                      {strings.NewTitle}
                                                    </strong>
                                                    {": " + title}
                                                  </span>
                                                </Field>
                                              }
                                              {
                                                startDate &&
                                                <Field style={{ display: "flex", alignItems: "flex-start" }}>
                                                  <CircleSmall20Regular style={{ minWidth: "20px" }} />
                                                  <span>
                                                    <strong>
                                                      {strings.NewStartDate}
                                                    </strong>
                                                    {
                                                      ": " +
                                                      /* @ts-ignore */
                                                      format(new Date(startDate), 'PPPPp', { locale: locales[locale] }).replace(/^\w/, (character: string): string => character.toUpperCase())
                                                    }
                                                  </span>
                                                </Field>
                                              }
                                              {
                                                endDate &&
                                                <Field style={{ display: "flex", alignItems: "flex-start" }}>
                                                  <CircleSmall20Regular style={{ minWidth: "20px" }} />
                                                  <span>
                                                    <strong>
                                                      {strings.NewEndDate}
                                                    </strong>
                                                    {
                                                      ": " +
                                                      /* @ts-ignore */
                                                      format(new Date(endDate), 'PPPPp', { locale: locales[locale] }).replace(/^\w/, (character: string): string => character.toUpperCase())
                                                    }
                                                  </span>
                                                </Field>
                                              }
                                              {
                                                meetingType &&
                                                <Field style={{ display: "flex", alignItems: "flex-start" }}>
                                                  <CircleSmall20Regular style={{ minWidth: "20px" }} />
                                                  <span>
                                                    <strong>
                                                      {strings.NewMeetingType}
                                                    </strong>
                                                    {": " + getTermLabel(attendanceTypes, meetingType)}
                                                  </span>
                                                </Field>
                                              }
                                              {
                                                location &&
                                                <Field style={{ display: "flex", alignItems: "flex-start" }}>
                                                  <CircleSmall20Regular style={{ minWidth: "20px" }} />
                                                  <span>
                                                    <strong>
                                                      {strings.NewLocation}
                                                    </strong>
                                                    {": " + location}
                                                  </span>
                                                </Field>
                                              }
                                              {
                                                locationInformation &&
                                                <Field style={{ display: "flex", alignItems: "flex-start" }}>
                                                  <CircleSmall20Regular style={{ minWidth: "20px" }} />
                                                  <span>
                                                    <strong>
                                                      {strings.NewLocationInformation}
                                                    </strong>
                                                    {": " + locationInformation}
                                                  </span>
                                                </Field>
                                              }
                                              {
                                                onlineTool &&
                                                <Field style={{ display: "flex", alignItems: "flex-start" }}>
                                                  <CircleSmall20Regular style={{ minWidth: "20px" }} />
                                                  <span>
                                                    <strong>
                                                      {strings.NewOnlineTool}
                                                    </strong>
                                                    {": " + getTermLabel(onlineTools, onlineTool)}
                                                  </span>
                                                </Field>
                                              }
                                              {
                                                onlineToolUrl &&
                                                <Field style={{ display: "flex", alignItems: "flex-start" }}>
                                                  <CircleSmall20Regular style={{ minWidth: "20px" }} />
                                                  <span>
                                                    <strong>
                                                      {strings.NewOnlineToolUrl}
                                                    </strong>
                                                    {": " + onlineToolUrl}
                                                  </span>
                                                </Field>
                                              }
                                              {
                                                description &&
                                                <Field style={{ display: "flex", alignItems: "flex-start" }}>
                                                  <CircleSmall20Regular style={{ minWidth: "20px" }} />
                                                  <span>
                                                    <strong>
                                                      {strings.NewDescription}
                                                    </strong>
                                                    {": " + description.replace(/(<([^>]+)>)/ig, '')}
                                                  </span>
                                                </Field>
                                              }
                                            </Field>
                                          </AccordionPanel>
                                        </AccordionItem>
                                      )
                                    }
                                    {
                                      /* Cambios en el orden del día */
                                      (agendaPoints && agendaPoints.length > 0) && (
                                        <AccordionItem value="2" >
                                          <AccordionHeader>
                                            {strings.MeetingPoints + ":"}
                                          </AccordionHeader>
                                          <AccordionPanel style={{ display: "flex", flexDirection: "column" }}>
                                            <Field style={{ paddingLeft: "27px" }}>
                                              <Field style={{ display: "flex", alignItems: "center" }}>
                                                <CircleSmall20Regular />
                                                {strings.TheOrderOfTheDayHasBeenUpdated}
                                              </Field>
                                            </Field>
                                          </AccordionPanel>
                                        </AccordionItem>
                                      )
                                    }
                                    {
                                      /* Adición de asistentes */
                                      (addedAssistants && addedAssistants.length > 0) && (
                                        <AccordionItem value="3">
                                          <AccordionHeader>
                                            {strings.AddedAssistants + ":"}
                                          </AccordionHeader>
                                          <AccordionPanel style={{ display: "flex", flexDirection: "column" }}>
                                            <Field style={{ paddingLeft: "27px" }}>
                                              {
                                                addedAssistants.map((addedAssistant: IPendingChanges): JSX.Element => (
                                                  <Field style={{ display: "flex", alignItems: "center" }}>
                                                    <CircleSmall20Regular />
                                                    {
                                                      <Person userId={addedAssistant.ChangeValue} view={ViewType.oneline} />
                                                    }
                                                  </Field>
                                                ))
                                              }
                                            </Field>
                                          </AccordionPanel>
                                        </AccordionItem>
                                      )
                                    }
                                    {
                                      /* Eliminación de asistentes */
                                      (removedAssistants && removedAssistants.length > 0) && (
                                        <AccordionItem value="4">
                                          <AccordionHeader>
                                            {strings.RemovedAssistants + ":"}
                                          </AccordionHeader>
                                          <AccordionPanel style={{ display: "flex", flexDirection: "column" }}>
                                            <Field style={{ paddingLeft: "27px" }}>
                                              {
                                                removedAssistants.map((removedAssistant: IPendingChanges): JSX.Element => (
                                                  <Field style={{ display: "flex", alignItems: "center" }}>
                                                    <CircleSmall20Regular />
                                                    {
                                                      <Person userId={removedAssistant.ChangeValue} view={ViewType.oneline} />
                                                    }
                                                  </Field>
                                                ))
                                              }
                                            </Field>
                                          </AccordionPanel>
                                        </AccordionItem>
                                      )
                                    }
                                    {
                                      /* Cambios en los Documents */
                                      (documents && documents.length > 0) && (
                                        <AccordionItem value="5">
                                          <AccordionHeader>
                                            {strings.Documents + ":"}
                                          </AccordionHeader>
                                          <AccordionPanel style={{ display: "flex", flexDirection: "column" }}>
                                            <Field style={{ paddingLeft: "27px" }}>
                                              {
                                                documents.map((document: IPendingChanges): JSX.Element => (
                                                  <Field style={{ display: "flex", alignItems: "center" }}>
                                                    <CircleSmall20Regular />
                                                    {document.ChangeValue}
                                                  </Field>
                                                ))
                                              }
                                            </Field>
                                          </AccordionPanel>
                                        </AccordionItem>
                                      )
                                    }
                                    {
                                      /* Cambios en los acuerdos */
                                      (agreements && agreements.length > 0) && (
                                        <AccordionItem value="6">
                                          <AccordionHeader>
                                            {strings.Agreements + ":"}
                                          </AccordionHeader>
                                          <AccordionPanel style={{ display: "flex", flexDirection: "column" }}>
                                            <Field style={{ paddingLeft: "27px" }}>
                                              {
                                                agreements.map((agreement: IPendingChanges): JSX.Element => (
                                                  <Field style={{ display: "flex", alignItems: "center" }}>
                                                    <CircleSmall20Regular />
                                                    {agreement.ChangeValue}
                                                  </Field>
                                                ))
                                              }
                                            </Field>
                                          </AccordionPanel>
                                        </AccordionItem>
                                      )
                                    }
                                    {
                                      /* Cambios en el Minutes */
                                      (minutes && minutes.length > 0) && (
                                        <AccordionItem value="7" >
                                          <AccordionHeader>
                                            {strings.Minutes + ":"}
                                          </AccordionHeader>
                                          <AccordionPanel style={{ display: "flex", flexDirection: "column" }}>
                                            <Field style={{ paddingLeft: "27px" }}>
                                              <Field style={{ display: "flex", alignItems: "center" }}>
                                                <CircleSmall20Regular />
                                                {strings.TheMinutesHasBeenUpdated}
                                              </Field>
                                            </Field>
                                          </AccordionPanel>
                                        </AccordionItem>
                                      )
                                    }
                                  </Accordion>
                                }
                              </div>
                              {
                                (index !== currentAlerts.length - 1) &&
                                <Divider style={{ marginTop: "5px", marginBottom: "5px" }} />
                              }
                            </>
                          );
                        })
                      }
                    </div>
                    :
                    /* Mensaje para indicar que no hay Alerts */
                    <div className={styles.popoverNoAlerts}>
                      <span> {strings.NoCurrentAlertsToShow} </span>
                    </div>
                }
              </PopoverSurface>
            </Popover>
          </span>
        </div>
      );

    }
  }

  private renderTheArchivedAlertsDrawer(): JSX.Element {
    const { locale } = this.props;
    const { loadingArchivedAlerts, archivedAlerts, drawerIsOpen, attendanceTypes, onlineTools, selectedArchivedAlert } = this.state;
    const locales = { "es-ES": es, "ca-ES": ca, "eu-ES": eu, "gl-ES": gl };
    return (
      <div>
        <Drawer
          type={'overlay'}
          size={'medium'}
          position={'end'}
          open={drawerIsOpen}
          onOpenChange={() => this.setState({ drawerIsOpen: false })}
        >
          <DrawerHeader>
            <DrawerHeaderTitle
              action={
                <Button
                  appearance="subtle"
                  icon={<Dismiss24Regular />}
                  onClick={(): void => this.setState({ drawerIsOpen: false })}
                />
              }
            >
              {
                /* Título del drawer */
                (archivedAlerts && archivedAlerts.length > 0) ?
                  strings.ArchivedAlerts + " (" + archivedAlerts.length + ")"
                  :
                  strings.ArchivedAlerts
              }
            </DrawerHeaderTitle>
          </DrawerHeader>
          <DrawerBody>
            <div className={styles.drawer}>
              {
                /* Se muestran los Alerts archivados */
                (!loadingArchivedAlerts) ?
                  (archivedAlerts && archivedAlerts.length > 0) ?
                    archivedAlerts.map((archivedAlert: INotification, index: number): JSX.Element => {

                      /* Filtrado de cambios en los datos básicos de la Meeting */
                      const title: string | undefined = this.filterPendingChanges("Title", archivedAlert.PendingChanges)?.[0]?.ChangeValue;
                      const location: string | undefined = this.filterPendingChanges("Location", archivedAlert.PendingChanges)?.[0]?.ChangeValue;
                      const locationInformation: string | undefined = this.filterPendingChanges("LocationDetails", archivedAlert.PendingChanges)?.[0]?.ChangeValue;
                      const meetingType: string | undefined = this.filterPendingChanges("MeetingType", archivedAlert.PendingChanges)?.[0]?.ChangeValue;
                      const description: string | undefined = this.filterPendingChanges("Description", archivedAlert.PendingChanges)?.[0]?.ChangeValue;
                      const startDate: string | undefined = this.filterPendingChanges("StartDate", archivedAlert.PendingChanges)?.[0]?.ChangeValue;
                      const endDate: string | undefined = this.filterPendingChanges("EndDate", archivedAlert.PendingChanges)?.[0]?.ChangeValue;
                      const onlineTool: string | undefined = this.filterPendingChanges("OnlineTool", archivedAlert.PendingChanges)?.[0]?.ChangeValue;
                      const onlineToolUrl: string | undefined = this.filterPendingChanges("UrlOnlineTool", archivedAlert.PendingChanges)?.[0]?.ChangeValue;

                      /* Filtrado de cambios en el orden del día */
                      const agendaPoints = this.filterPendingChanges("AgendaItem", archivedAlert.PendingChanges);

                      /* Filtrado de cambios en los asistentes */
                      const addedAssistants: IPendingChanges[] | undefined = this.filterPendingChanges("AddAttendance", archivedAlert.PendingChanges);
                      const removedAssistants: IPendingChanges[] | undefined = this.filterPendingChanges("RemoveAttendance", archivedAlert.PendingChanges);

                      /* Filtrado de cambios en los Documents */
                      const documents: IPendingChanges[] | undefined = this.filterPendingChanges("Document", archivedAlert.PendingChanges);

                      /* Filtrado de cambios en los acuerdos */
                      const agreements: IPendingChanges[] | undefined = this.filterPendingChanges("Agreement", archivedAlert.PendingChanges);

                      /* Filtrado de cambios en el Minutes */
                      const minutes: IPendingChanges[] | undefined = this.filterPendingChanges("Minutes", archivedAlert.PendingChanges);
                      return (
                        <>
                          <div className={styles.drawerRow}>
                            <div className={styles.drawerRowTop}>
                              {
                                (archivedAlert.NotificationType === TipoAlert.Sistema) ?
                                  <TextBulletListSquareWarning20Regular
                                    className={styles.alertIcon}
                                    style={{ minHeight: "20px", minWidth: "20px" }}
                                  />
                                  :
                                  <AlertOn20Regular
                                    className={styles.alertIcon}
                                    style={{ minHeight: "20px", minWidth: "20px" }}
                                  />
                              }
                              <span
                                style={{ cursor: "pointer", width: "100%" }}
                                onClick={(): void => this.onClickAlert(archivedAlert)}
                                title={strings.GoToTheMeeting + archivedAlert.Title}
                              >
                                {archivedAlert.Title}
                              </span>
                            </div>
                            <div className={styles.drawerRowBottom}>
                              <span
                                style={{ cursor: "pointer", width: "100%" }}
                                onClick={(): void => this.onClickAlert(archivedAlert)}
                                title={strings.GoToTheMeeting + archivedAlert.Title}
                              >
                                {archivedAlert.Body}
                              </span>
                              {
                                (archivedAlert.PendingChanges && archivedAlert.PendingChanges.length > 0) ?
                                  (selectedArchivedAlert === archivedAlert) ?
                                    <span>
                                      <Link
                                        style={{ width: "fit-content" }}
                                        onClick={(): void => this.setState({ selectedArchivedAlert: undefined })}
                                      >
                                        {strings.SeeLess}
                                      </Link>.
                                    </span>
                                    :
                                    <span>
                                      <Link
                                        style={{ width: "fit-content" }}
                                        onClick={(): void => this.setState({ selectedArchivedAlert: archivedAlert })}
                                      >
                                        {strings.SeeMore}
                                      </Link>.
                                    </span>
                                  :
                                  undefined
                              }
                            </div>
                            {
                              /* Desplegable de un Alert asociado a una Meeting actualizada */
                              (selectedArchivedAlert === archivedAlert) &&
                              <Accordion collapsible style={{ display: "flex", flexDirection: "column", paddingLeft: "13px" }}>
                                {
                                  /* Cambios en los datos básicos de la Meeting */
                                  (title || location || locationInformation || meetingType || description || startDate || endDate || onlineTool || onlineToolUrl) && (
                                    <AccordionItem value="1">
                                      <AccordionHeader>
                                        <Field>
                                          <span>
                                            {strings.BasicData + ":"}
                                          </span>
                                        </Field>
                                      </AccordionHeader>
                                      <AccordionPanel style={{ display: "flex", flexDirection: "column" }}>
                                        <Field style={{ paddingLeft: "27px" }}>
                                          {
                                            title &&
                                            <Field style={{ display: "flex", alignItems: "flex-start" }}>
                                              <CircleSmall20Regular style={{ minWidth: "20px" }} />
                                              <span>
                                                <strong>
                                                  {strings.NewTitle}
                                                </strong>
                                                {": " + title}
                                              </span>
                                            </Field>
                                          }
                                          {
                                            startDate &&
                                            <Field style={{ display: "flex", alignItems: "flex-start" }}>
                                              <CircleSmall20Regular style={{ minWidth: "20px" }} />
                                              <span>
                                                <strong>
                                                  {strings.NewStartDate}
                                                </strong>
                                                {
                                                  ": " +
                                                  /* @ts-ignore */
                                                  format(new Date(startDate), 'PPPPp', { locale: locales[locale] }).replace(/^\w/, (character: string): string => character.toUpperCase())
                                                }
                                              </span>
                                            </Field>
                                          }
                                          {
                                            endDate &&
                                            <Field style={{ display: "flex", alignItems: "flex-start" }}>
                                              <CircleSmall20Regular style={{ minWidth: "20px" }} />
                                              <span>
                                                <strong>
                                                  {strings.NewEndDate}
                                                </strong>
                                                {
                                                  ": " +
                                                  /* @ts-ignore */
                                                  format(new Date(endDate), 'PPPPp', { locale: locales[locale] }).replace(/^\w/, (character: string): string => character.toUpperCase())
                                                }
                                              </span>
                                            </Field>
                                          }
                                          {
                                            meetingType &&
                                            <Field style={{ display: "flex", alignItems: "flex-start" }}>
                                              <CircleSmall20Regular style={{ minWidth: "20px" }} />
                                              <span>
                                                <strong>
                                                  {strings.NewMeetingType}
                                                </strong>
                                                {": " + getTermLabel(attendanceTypes, meetingType)}
                                              </span>
                                            </Field>
                                          }
                                          {
                                            location &&
                                            <Field style={{ display: "flex", alignItems: "flex-start" }}>
                                              <CircleSmall20Regular style={{ minWidth: "20px" }} />
                                              <span>
                                                <strong>
                                                  {strings.NewLocation}
                                                </strong>
                                                {": " + location}
                                              </span>
                                            </Field>
                                          }
                                          {
                                            locationInformation &&
                                            <Field style={{ display: "flex", alignItems: "flex-start" }}>
                                              <CircleSmall20Regular style={{ minWidth: "20px" }} />
                                              <span>
                                                <strong>
                                                  {strings.NewLocationInformation}
                                                </strong>
                                                {": " + locationInformation}
                                              </span>
                                            </Field>
                                          }
                                          {
                                            onlineTool &&
                                            <Field style={{ display: "flex", alignItems: "flex-start" }}>
                                              <CircleSmall20Regular style={{ minWidth: "20px" }} />
                                              <span>
                                                <strong>
                                                  {strings.NewOnlineTool}
                                                </strong>
                                                {": " + getTermLabel(onlineTools, onlineTool)}
                                              </span>
                                            </Field>
                                          }
                                          {
                                            onlineToolUrl &&
                                            <Field style={{ display: "flex", alignItems: "flex-start" }}>
                                              <CircleSmall20Regular style={{ minWidth: "20px" }} />
                                              <span>
                                                <strong>
                                                  {strings.NewOnlineToolUrl}
                                                </strong>
                                                {": " + onlineToolUrl}
                                              </span>
                                            </Field>
                                          }
                                          {
                                            description &&
                                            <Field style={{ display: "flex", alignItems: "flex-start" }}>
                                              <CircleSmall20Regular style={{ minWidth: "20px" }} />
                                              <span>
                                                <strong>
                                                  {strings.NewDescription}
                                                </strong>
                                                {": " + description.replace(/(<([^>]+)>)/ig, '')}
                                              </span>
                                            </Field>
                                          }
                                        </Field>
                                      </AccordionPanel>
                                    </AccordionItem>
                                  )
                                }
                                {
                                  /* Cambios en el orden del día */
                                  (agendaPoints && agendaPoints.length > 0) && (
                                    <AccordionItem value="2" >
                                      <AccordionHeader>
                                        {strings.MeetingPoints + ":"}
                                      </AccordionHeader>
                                      <AccordionPanel style={{ display: "flex", flexDirection: "column" }}>
                                        <Field style={{ paddingLeft: "27px" }}>
                                          <Field style={{ display: "flex", alignItems: "center" }}>
                                            <CircleSmall20Regular />
                                            {strings.TheOrderOfTheDayHasBeenUpdated}
                                          </Field>
                                        </Field>
                                      </AccordionPanel>
                                    </AccordionItem>
                                  )
                                }
                                {
                                  /* Adición de asistentes */
                                  (addedAssistants && addedAssistants.length > 0) && (
                                    <AccordionItem value="3">
                                      <AccordionHeader>
                                        {strings.AddedAssistants + ":"}
                                      </AccordionHeader>
                                      <AccordionPanel style={{ display: "flex", flexDirection: "column" }}>
                                        <Field style={{ paddingLeft: "27px" }}>
                                          {
                                            addedAssistants.map((addedAssistant: IPendingChanges): JSX.Element => (
                                              <Field style={{ display: "flex", alignItems: "center" }}>
                                                <CircleSmall20Regular />
                                                {
                                                  <Person userId={addedAssistant.ChangeValue} view={ViewType.oneline} />
                                                }
                                              </Field>
                                            ))
                                          }
                                        </Field>
                                      </AccordionPanel>
                                    </AccordionItem>
                                  )
                                }
                                {
                                  /* Eliminación de asistentes */
                                  (removedAssistants && removedAssistants.length > 0) && (
                                    <AccordionItem value="4">
                                      <AccordionHeader>
                                        {strings.RemovedAssistants + ":"}
                                      </AccordionHeader>
                                      <AccordionPanel style={{ display: "flex", flexDirection: "column" }}>
                                        <Field style={{ paddingLeft: "27px" }}>
                                          {
                                            removedAssistants.map((removedAssistant: IPendingChanges): JSX.Element => (
                                              <Field style={{ display: "flex", alignItems: "center" }}>
                                                <CircleSmall20Regular />
                                                {
                                                  <Person userId={removedAssistant.ChangeValue} view={ViewType.oneline} />
                                                }
                                              </Field>
                                            ))
                                          }
                                        </Field>
                                      </AccordionPanel>
                                    </AccordionItem>
                                  )
                                }
                                {
                                  /* Cambios en los Documents */
                                  (documents && documents.length > 0) && (
                                    <AccordionItem value="5">
                                      <AccordionHeader>
                                        {strings.Documents + ":"}
                                      </AccordionHeader>
                                      <AccordionPanel style={{ display: "flex", flexDirection: "column" }}>
                                        <Field style={{ paddingLeft: "27px" }}>
                                          {
                                            documents.map((document: IPendingChanges): JSX.Element => (
                                              <Field style={{ display: "flex", alignItems: "center" }}>
                                                <CircleSmall20Regular />
                                                {document.ChangeValue}
                                              </Field>
                                            ))
                                          }
                                        </Field>
                                      </AccordionPanel>
                                    </AccordionItem>
                                  )
                                }
                                {
                                  /* Cambios en los acuerdos */
                                  (agreements && agreements.length > 0) && (
                                    <AccordionItem value="6">
                                      <AccordionHeader>
                                        {strings.Agreements + ":"}
                                      </AccordionHeader>
                                      <AccordionPanel style={{ display: "flex", flexDirection: "column" }}>
                                        <Field style={{ paddingLeft: "27px" }}>
                                          {
                                            agreements.map((agreement: IPendingChanges): JSX.Element => (
                                              <Field style={{ display: "flex", alignItems: "center" }}>
                                                <CircleSmall20Regular />
                                                {agreement.ChangeValue}
                                              </Field>
                                            ))
                                          }
                                        </Field>
                                      </AccordionPanel>
                                    </AccordionItem>
                                  )
                                }
                                {
                                  /* Cambios en el Minutes */
                                  (minutes && minutes.length > 0) && (
                                    <AccordionItem value="7" >
                                      <AccordionHeader>
                                        {strings.Minutes + ":"}
                                      </AccordionHeader>
                                      <AccordionPanel style={{ display: "flex", flexDirection: "column" }}>
                                        <Field style={{ paddingLeft: "27px" }}>
                                          <Field style={{ display: "flex", alignItems: "center" }}>
                                            <CircleSmall20Regular />
                                            {strings.TheMinutesHasBeenUpdated}
                                          </Field>
                                        </Field>
                                      </AccordionPanel>
                                    </AccordionItem>
                                  )
                                }
                              </Accordion>
                            }

                          </div>
                          {
                            (index !== archivedAlerts.length - 1) &&
                            <Divider style={{ marginTop: "5px", marginBottom: "5px" }} />
                          }
                        </>
                      );
                    })
                    :
                    <div className={styles.drawerThereAreNoAlerts}>
                      <span> {strings.NoArchivedAlertsToShow} </span>
                    </div>
                  :
                  <Spinner
                    style={{ padding: "10px 0", display: "flex", justifyContent: "flex-start" }}
                    label={strings.LoadingArchivedAlerts + "..."}
                  />
              }
            </div>
          </DrawerBody>
        </Drawer>
      </div>
    );
  }

  public render(): React.ReactElement<IWelcomeMessageprops> {
    const { context } = this.props;
    const { dialogIsOpen, loadingUserInformation, savingUserInformation, backendHasFailed } = this.state;
    return (
      <section className={styles.welcomeMessage}>
        <div className={styles.message}>
          <Text as='h1' className={styles.hiMessage}>
            {strings.HiMessage}, {context.pageContext.user.displayName}
          </Text>
          <Edit24Regular className={styles.iconButton} onClick={this.openDialog} />
        </div>
        <div>
          {
            /* Barra de Alerts */
            this.renderAlerts()
          }
        </div>
        <Dialog open={dialogIsOpen} onOpenChange={(ev, data): void => this.setState({ dialogIsOpen: data.open })}>
          <DialogSurface>
            <DialogBody>
              <DialogTitle> {strings.PersonalConfig} </DialogTitle>
              <DialogContent>
                {
                  loadingUserInformation ?
                    this.onRenderSkeleton()
                    :
                    this.onRenderForm()
                }
              </DialogContent>
              <DialogActions className={styles.footer}>
                {
                  savingUserInformation &&
                  <Spinner size='tiny' label={strings.Saving + "..."} />
                }
                <Button appearance='secondary' onClick={this.closeDialog}>
                  {strings.Discard}
                </Button>
                <Button disabled={loadingUserInformation || savingUserInformation || backendHasFailed} form='details' type='submit' appearance='primary'>
                  {strings.Save}
                </Button>
              </DialogActions>
            </DialogBody>
          </DialogSurface>
        </Dialog>
      </section>
    );
  }

}