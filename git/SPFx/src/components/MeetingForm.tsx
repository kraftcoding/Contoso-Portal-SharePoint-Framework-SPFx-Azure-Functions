import * as React from 'react';
import * as yup from 'yup';
import * as strings from 'MeetingsWebPartStrings';
import styles from './MeetingForm.module.scss';
import { IMeetingFormState } from './IMeetingFormState';
import { IMeetingFormprops } from './IMeetingFormprops';
import {
    ArrowRight20Regular,
    Calendar24Filled,
    Clock20Regular,
    DesktopToolbox20Regular,
    Edit20Regular,
    Link20Regular,
    Location20Regular,
    NotepadEdit20Regular,
    Options20Regular,
    PeopleAdd20Regular,
    PersonAdd20Regular,
    TextAlignLeft20Regular
} from '@fluentui/react-icons';
import {
    Dialog,
    DialogBody, DialogContent,
    DialogSurface,
    DialogTitle,
    DialogTrigger,
    Dropdown,
    Input,
    Option,
    Button,
    Field,
    DialogActions,
    Spinner
} from '@fluentui/react-components';
import { PeoplePicker } from '@microsoft/mgt-react/dist/es6/spfx';
import format from 'date-fns/format';
import ReactQuill, { Quill } from 'react-quill';
import 'react-quill/dist/quill.snow.css';
import { Term } from '../models/ITag';
import { ITermInfo } from '@pnp/sp/taxonomy';
import { Form, Formik } from 'formik';
import { getTermLabel } from '../utils/Utils';
import { getBodyIdFormUrl } from '../service/RoleService';
import { PersonType } from '@microsoft/mgt-spfx';
import { EventStatus, IEvent, IUserAttendance } from '../service/BackendServiceModels/EventModels';
import { IMember } from '../service/BackendServiceModels/MemberModel';
import { UserRoles } from '../models/IUserInfo';

const ReuneteToolTermId: string = '8f9f34b6-307d-4361-8957-bfff9df1d1a9';
const OnlineTermId: string = '8f4e0e22-2c3f-4454-8210-f4151dceeb89';
const OnlineInPersonTermId: string = 'e49c8fe8-87f5-4b62-b6ce-0df240b7ec7f';
const InPersonTermId: string = '3a16b292-5616-4ef9-8a56-2b36c20bb992';
// const WrittenprodcedureTermId: string = '474ce74e-26c2-4ada-9dbb-516be5deb394';
const OnlineToolSetId: string = '55dfa3e7-f0df-4dc8-b53c-4018418041fe';
const AtendanceTypeSetId: string = '0775b014-e8fc-4aac-ae49-6b0b901156c1';

const BaseLink = Quill.import("formats/link");
class Link extends BaseLink {
    format(name: any, value: any) {
        if (["href", "target"].indexOf(name) > -1) {
            if (value) {
                this.domNode.setAttribute(name, value);
            } else {
                this.domNode.removeAttribute(name);
            }
        } else {
            super.format(name, value);
        }
    }
    static create(value: any) {
        const node = super.create(value);
        let preview = this.sanitize(value);
        if (!/^(http|https|tel|mailto):/.test(preview)) {
            // Change HERE: prefix with https if none of the expected schemes is testsent
            preview = "https://" + preview
        }

        node.setAttribute("href", preview);
        node.setAttribute("target", "_blank");
        return node;
    }
}
Quill.register(Link, true);

export default class MeetingForm extends React.Component<IMeetingFormprops, IMeetingFormState> {

    constructor(props: IMeetingFormprops) {
        super(props);
        this.state = {
            onlineTools: [],
            attendanceTypes: [],
            eventToolId: this.props.eventData ? this.props.eventData.EventToolId : "",
            keepTheFormOpen: true,
            submitting: false,
            attendanceTypeSelected: this.props.eventData ? this.props.eventData.AttendanceTypeId : "",
            members: [],
            formDataLoading: true,
            membersGroupId: "",
            guestsGroupId: "",
            convoGroupId: "",
            gestorGroupId: "",
            asistenteGroupId: "",
            initialFormFieldValues: {
                Title: this.props.eventData?.Title || "",
                Description: this.props.eventData?.Description || "",
                StartDate: this.props.eventData?.StartDate || new Date(),
                EndDate: this.props.eventData?.EndDate || new Date(),
                Location: this.props.eventData?.Location || "",
                LocationDetails: this.props.eventData?.LocationDetails || "",
                EventToolId: this.props.eventData?.EventToolId || "",
                MeetingToolUrl: this.props.eventData?.MeetingToolUrl || "",
                AttendanceTypeId: this.props.eventData?.AttendanceTypeId || "",
                Attendees: [],
                Guests: []
            }
        }
    }

    public async componentDidMount(): promise<void> {
        await this.onInit();
    }

    private async onInit(): promise<void> {
        const { spService, context, bkService, eventData } = this.props;
        const { initialFormFieldValues } = this.state;

        const bodyId = getBodyIdFormUrl(context.pageContext.site.serverRelativeUrl);
        const isArchived: boolean = eventData?.StatusId === EventStatus.Archived;
        let membersGroupId: string = "";
        let guestsGroupId: string = "";
        let convoGroupId: string = "";
        let gestorGroupId: string = "";
        let asistenteGroupId: string = "";
        let onlineToolRes: ITermInfo[] = [];
        let attendaceRes: ITermInfo[] = [];

        let members: IUserAttendance[] | IMember[] = [];
        const attendees: string[] = [];
        const guests: string[] = [];

        try {
            [
                membersGroupId,
                guestsGroupId,
                convoGroupId,
                gestorGroupId,
                asistenteGroupId,
                onlineToolRes,
                attendaceRes
            ] = await promise.all([
                await spService.getGroupId(`${bodyId}_members`),
                await spService.getGroupId(`${bodyId}_guests`),
                await spService.getGroupId(`${bodyId}_schedulers`),
                await spService.getGroupId(`${bodyId}_gestorschedulers`),
                await spService.getGroupId(`${bodyId}_asistentemembers`),
                await spService.getTaxonomy(OnlineToolSetId),
                await spService.getTaxonomy(AtendanceTypeSetId)
            ]);
        }
        catch (error) {
            console.error(error);
        }

        try {
            if (eventData) {
                /* En una Meeting ya existente, se testcargan los usuarios asociados a la misma */
                members = await bkService.getAttendanceByEventId(eventData.BodyId, eventData.Id, isArchived);
            }
            else {
                /* En la creación de una Meeting, se testcargan los members del órgano */
                members = await bkService.getBodyUsersByRole(context.pageContext.web.serverRelativeUrl.split("sites/")[1], UserRoles.Members);
            }
        }
        catch (error) {
            console.error(error);
        }

        if (members) {
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            members.forEach((item: any): void => {
                if (item.IsGuest) {
                    guests.push(item.UserPrincipalName)
                }
                else {
                    attendees.push(item.UserPrincipalName);
                }
            });
        }

        const onlineTools: Term[] = onlineToolRes.map((tag: ITermInfo) => Term.mapSPToTerm(tag, context.pageContext.cultureInfo.currentUICultureName));
        const attendanceTypes: Term[] = attendaceRes.map((tag: ITermInfo) => Term.mapSPToTerm(tag, context.pageContext.cultureInfo.currentUICultureName));

        this.setState({
            formDataLoading: false,
            membersGroupId,
            guestsGroupId,
            convoGroupId,
            gestorGroupId,
            asistenteGroupId,
            onlineTools,
            attendanceTypes,
            initialFormFieldValues: {
                ...initialFormFieldValues,
                Attendees: attendees,
                Guests: guests
            }
        });
    }

    private async onSave(values: any): promise<void> {
        const { context, bkService, editMode, eventData, openForm, dataChanged } = this.props;
        try {
            this.setState({ submitting: true });
            if (editMode && eventData) {
                console.log(values);
                const refreshedItem: IEvent = await bkService.updateEvent(context.pageContext.web.serverRelativeUrl.split("sites/")[1], eventData?.Id, values);
                dataChanged(refreshedItem);
            }
            else {
                const createdEvent: IEvent = await bkService.createEvent(context.pageContext.web.serverRelativeUrl.split("sites/")[1], values);
                history.pushState({}, "", context.pageContext.web.absoluteUrl + "#/" + createdEvent.Id); // Se hace un redirect al evento creado
                window.location.reload();
            }
            this.setState({ submitting: true, keepTheFormOpen: false });
        }
        catch (error) {
            this.setState({ keepTheFormOpen: true, submitting: false });
            console.log(error);
        }
        finally {
            openForm();
        }
    }

    public render(): React.ReactElement<IMeetingFormprops> {
        const { isEditor, editMode, readOnly } = this.props;
        const {
            onlineTools,
            attendanceTypes,
            attendanceTypeSelected,
            eventToolId,
            submitting,
            initialFormFieldValues,
            formDataLoading,
            membersGroupId,
            guestsGroupId,
            convoGroupId,
            gestorGroupId,
            asistenteGroupId,
            keepTheFormOpen
        } = this.state;
        const locked: boolean = !isEditor || readOnly;

        let langStyle = styles"en-US";
        switch (this.props.context.pageContext.cultureInfo.currentUICultureName) {
            case 'ca-ES':
                langStyle = styles.caES;
                break;
            case 'es-ES':
                langStyle = styles"en-US";
                break;
            case 'eu-ES':
                langStyle = styles.euES;
                break;
            case 'gl-ES':
                langStyle = styles.glES;
                break;
            default:
                break;
        }

        const validationSchema = yup.object().shape({
            Title: yup
                .string()
                .min(2, strings.MinimumCharactersForTitle)
                .max(255, strings.MaximumCharactersForTitle)
                .required(strings.RequiredField)
                .matches(/^(.*)?\S+(.*)?$/, strings.RequiredField),
            Description: yup
                .string(),
            StartDate: yup
                .date()
                .required(strings.RequiredField)
                .min(initialFormFieldValues.StartDate < new Date() ? initialFormFieldValues.StartDate : new Date(), strings.StartDateLessThanTheCurrentDate),
            EndDate: yup
                .date()
                .required(strings.RequiredField)
                .min(yup.ref('StartDate'), strings.EndDateLessThanStartDate),
            AttendanceTypeId: yup
                .string()
                .required(strings.RequiredAttendanceTypeId),
            Location: yup
                .string()
                .when('AttendanceTypeId', ([AttendanceTypeId], schema) => {
                    if (AttendanceTypeId === InPersonTermId || AttendanceTypeId === OnlineInPersonTermId)
                        return yup
                            .string()
                            .trim()
                            .required(strings.RequiredField);
                    return schema;
                })
                .max(255, strings.MaximumCharactersForLocation),
            LocationDetails: yup
                .string()
                .max(255, strings.MaximumCharactersForLocationDetails),
            EventToolId: yup
                .string()
                .when('AttendanceTypeId', ([AttendanceTypeId], schema) => {
                    if (AttendanceTypeId === OnlineTermId || AttendanceTypeId === OnlineInPersonTermId)
                        return yup
                            .string()
                            .trim()
                            .required(strings.RequiredField);
                    return schema;
                }),
            MeetingToolUrl: yup
                .string()
                .when('EventToolId', ([EventToolId], schema) => {
                    if (EventToolId === ReuneteToolTermId)
                        return yup
                            .string()
                            .trim()
                            .required(strings.RequiredUrl);
                    return schema;
                })
                .matches(/((https?):\/\/)?(www.)?.*$/, strings.RequiredValidUrl)
                .max(255, strings.MaximumCharactersForUrl),
            Attendees: yup
                .array().of(yup.string())
                .required(strings.RequiredAttendees)
                .min(1, strings.RequiredAttendees)
        });
        return (
            <Dialog open={keepTheFormOpen}>
                <DialogSurface className={styles.dialogSurface}>
                    <DialogBody>
                        <DialogTitle className={styles.dialogTitle}>
                            <Calendar24Filled />
                            <span>
                                {
                                    !locked ?
                                        editMode ?
                                            strings.EditMeeting
                                            :
                                            strings.NewMeeting
                                        :
                                        strings.MeetingDetails
                                }
                            </span>
                        </DialogTitle>
                        <DialogContent>
                            {!formDataLoading ?
                                <Formik enableReinitialize initialValues={initialFormFieldValues} validationSchema={validationSchema}
                                    validateOnChange={false} validateOnBlur={false} onSubmit={(values): promise<void> => this.onSave(values)}>
                                    {
                                        (formik: any) => {
                                            const {
                                                values,
                                                handleChange,
                                                handleSubmit,
                                                initialValues,
                                                errors,
                                                setFieldValue
                                            } = formik;
                                            return (<Form id="convo" className={styles.form} noValidate onSubmit={handleSubmit}>
                                                {/* Title */}
                                                <div className={styles.formRow}>
                                                    <Edit20Regular className={styles.icon} />
                                                    <div className={styles.formItem}>
                                                        <Field validationMessage={errors.Title}>
                                                            <Input
                                                                name='Title'
                                                                type='text'
                                                                placeholder={strings.TitlePlaceholder}
                                                                minLength={1}
                                                                maxLength={255}
                                                                value={values.Title}
                                                                onChange={handleChange}
                                                                disabled={locked}
                                                            />
                                                        </Field>
                                                    </div>
                                                </div>
                                                <div className={styles.formRow} >
                                                    {/* Start date */}
                                                    <div className={styles.formRow} >
                                                        <Clock20Regular className={styles.icon} />
                                                        <div className={styles.formItem}>
                                                            <Field validationMessage={errors.StartDate}>
                                                                <Input
                                                                    key='StartDate'
                                                                    value={values.StartDate && format(new Date(values.StartDate), 'yyyy-MM-dd\'T\'HH:mm')}
                                                                    type='datetime-local'
                                                                    onChange={async (data): promise<void> => {
                                                                        await setFieldValue("StartDate", new Date(data.target.value).toISOString());
                                                                        if (new Date(values.EndDate) < new Date(data.target.value)) {
                                                                            await setFieldValue("EndDate", new Date(data.target.value).toISOString());
                                                                        }
                                                                    }}
                                                                    //onKeyDown={(e) => e.preventDefault()}
                                                                    disabled={locked}
                                                                />
                                                            </Field>
                                                        </div>
                                                    </div>
                                                    {/* End date */}
                                                    <div className={styles.formRow}>
                                                        <ArrowRight20Regular className={`${styles.icon} ${styles.spaceLeft}`} />
                                                        <div className={styles.formItem} >
                                                            <Field validationMessage={errors.EndDate}>
                                                                <Input
                                                                    key='EndDate'
                                                                    value={values.EndDate && format(new Date(values.EndDate), 'yyyy-MM-dd\'T\'HH:mm')}
                                                                    type='datetime-local'
                                                                    onChange={async (data): promise<void> => {
                                                                        await setFieldValue("EndDate", new Date(data.target.value).toISOString());
                                                                    }}
                                                                    //onKeyDown={(e): void => e.preventDefault()}
                                                                    disabled={locked}
                                                                />
                                                            </Field>
                                                        </div>
                                                    </div>
                                                </div>
                                                <div className={styles.formRow}>
                                                    {/* AttendanceTypeId */}
                                                    <div className={styles.formRow}>
                                                        <Options20Regular className={styles.icon} />
                                                        <div className={styles.formItem} >
                                                            <Field validationMessage={errors.AttendanceTypeId}>
                                                                <Dropdown
                                                                    name='AttendanceTypeId'
                                                                    placeholder={strings.AttendanceTypePlaceholder}
                                                                    value={getTermLabel(this.state.attendanceTypes, values.AttendanceTypeId)}
                                                                    onOptionSelect={async (e, data) => {
                                                                        if (data.optionValue !== OnlineTermId && data.optionValue !== OnlineInPersonTermId) {
                                                                            await setFieldValue("EventToolId", "");
                                                                            await setFieldValue("MeetingToolUrl", "");
                                                                            this.setState({ eventToolId: "" });
                                                                        }
                                                                        await setFieldValue("AttendanceTypeId", data.optionValue);
                                                                        this.setState({ attendanceTypeSelected: data.optionValue || "" });
                                                                    }}
                                                                    disabled={locked}
                                                                >
                                                                    {
                                                                        attendanceTypes.map((tool: Term): JSX.Element =>
                                                                            <Option value={tool.key} key={tool.key}>{tool.text}</Option>)
                                                                    }
                                                                </Dropdown>
                                                            </Field>
                                                        </div>
                                                    </div>
                                                    {/* EventToolId */}
                                                    {
                                                        (attendanceTypeSelected === OnlineTermId || attendanceTypeSelected === OnlineInPersonTermId) &&
                                                        <div className={styles.formRow} >
                                                            <DesktopToolbox20Regular className={`${styles.icon} ${styles.spaceLeft}`} />
                                                            <div className={styles.formItem}>
                                                                <Field validationMessage={errors.EventToolId}>
                                                                    <Dropdown
                                                                        name='EventToolId'
                                                                        placeholder={strings.OnlineToolPlaceholder}
                                                                        selectedOptions={[this.state.eventToolId]}
                                                                        value={getTermLabel(this.state.onlineTools, values.EventToolId)}
                                                                        onOptionSelect={async (e, data) => {
                                                                            /*if (data.optionValue !== reuneteToolTermId) {
                                                                                await setFieldValue("MeetingToolUrl", "");
                                                                            }*/
                                                                            await setFieldValue("EventToolId", data.optionValue);
                                                                            this.setState({ eventToolId: data.optionValue || "" });
                                                                        }}
                                                                        disabled={locked}
                                                                    >
                                                                        {
                                                                            onlineTools.map((tool: Term): JSX.Element =>
                                                                                <Option value={tool.key} key={tool.key}>{tool.text}</Option>)
                                                                        }
                                                                    </Dropdown>
                                                                </Field>
                                                            </div>
                                                        </div>
                                                    }
                                                </div>
                                                {/* Location */}
                                                <div className={styles.formRow}>
                                                    {
                                                        (attendanceTypeSelected === InPersonTermId || attendanceTypeSelected === OnlineInPersonTermId) &&
                                                        <div className={styles.formRow}>
                                                            <Location20Regular className={styles.icon} />
                                                            <div className={styles.formItem} >
                                                                <Field validationMessage={errors.Location}>
                                                                    <Input
                                                                        name='Location'
                                                                        placeholder={strings.LocationPlaceholder}
                                                                        type='text'
                                                                        value={values.Location}
                                                                        onChange={handleChange}
                                                                        disabled={locked}
                                                                    />
                                                                </Field>
                                                            </div>
                                                        </div>
                                                    }
                                                    {
                                                        eventToolId === ReuneteToolTermId &&
                                                        <div className={styles.formRow} >
                                                            <Link20Regular className={`${styles.icon} ${styles.spaceLeft}`} />
                                                            <div className={styles.formItem}>
                                                                <Field validationMessage={errors.MeetingToolUrl}>
                                                                    <Input
                                                                        placeholder={strings.UrlPlaceholder}
                                                                        type='url'
                                                                        name='MeetingToolUrl'
                                                                        onChange={handleChange}
                                                                        value={values.MeetingToolUrl}
                                                                        disabled={locked}
                                                                    />
                                                                </Field>
                                                            </div>
                                                        </div>
                                                    }
                                                </div>
                                                {
                                                    (attendanceTypeSelected === InPersonTermId || attendanceTypeSelected === OnlineInPersonTermId) &&
                                                    <div className={`${styles.formRow}`} >
                                                        <TextAlignLeft20Regular className={styles.icon} />
                                                        <div className={styles.formItem}>
                                                            <Field validationMessage={errors.LocationDetails}>
                                                                <Input
                                                                    placeholder={strings.LocationDescriptionPlaceholder}
                                                                    type='text'
                                                                    name='LocationDetails'
                                                                    onChange={handleChange}
                                                                    value={values.LocationDetails}
                                                                    disabled={locked}
                                                                />
                                                            </Field>
                                                        </div>
                                                    </div>
                                                }
                                                <div className={styles.formRow}>
                                                    <PersonAdd20Regular className={styles.icon} />
                                                    <div className={styles.formItem} >
                                                        <Field validationMessage={errors.Attendees}>
                                                            <PeoplePicker
                                                                key={"Attendees"}
                                                                defaultSelectedUserIds={initialValues.Attendees}
                                                                type={PersonType.person}

                                                                groupIds={[membersGroupId, convoGroupId]}
                                                                placeholder={strings.PeoplePlaceholder}
                                                                transitiveSearch={true}
                                                                selectionChanged={async (e) => {
                                                                    await setFieldValue("Attendees", e.detail.map((value: any, index, array) => value.userPrincipalName));
                                                                }}
                                                                disabled={locked}
                                                            >
                                                            </PeoplePicker>
                                                        </Field>
                                                    </div>
                                                </div>
                                                <div className={styles.formRow}>
                                                    <PeopleAdd20Regular className={styles.icon} />
                                                    <div className={styles.formItem} >
                                                        <Field validationMessage={errors.Guests}>
                                                            <PeoplePicker
                                                                key={"Guests"}
                                                                defaultSelectedUserIds={initialValues.Guests}
                                                                placeholder={strings.GuestPeoplePlaceholder}
                                                                type={PersonType.person}

                                                                groupIds={[guestsGroupId, asistenteGroupId, gestorGroupId]}
                                                                //groupIds={[guestsGroupId, gestorGroupId]}
                                                                transitiveSearch={true}
                                                                selectionChanged={async (e) => {
                                                                    await setFieldValue("Guests", e.detail.map((value: any, index, array) => value.userPrincipalName));
                                                                }}
                                                                disabled={locked}
                                                            />
                                                        </Field>
                                                    </div>
                                                </div>
                                                {/* Description */}
                                                <div className={styles.formRow}>
                                                    <NotepadEdit20Regular className={styles.icon} />
                                                    <div className={styles.formItem}>
                                                        <Field validationMessage={errors.Description}>
                                                            <ReactQuill
                                                                key={"Description"}
                                                                placeholder={strings.DescriptionPlaceholder}
                                                                readOnly={locked}
                                                                className={`${styles.richText} ${langStyle}`}
                                                                modules={{
                                                                    toolbar: !locked ? [
                                                                        ['bold', 'italic', 'underline'],
                                                                        [{ 'list': 'ordered' }, { 'list': 'bullet' }, { 'indent': '-1' }, { 'indent': '+1' }],
                                                                        ['link']
                                                                    ] : false
                                                                }}
                                                                formats={!locked ? [
                                                                    'bold', 'italic', 'underline',
                                                                    'list', 'bullet', 'indent',
                                                                    'link'
                                                                ] : undefined}
                                                                theme="snow"
                                                                value={values.Description}
                                                                onChange={(text: any) => {
                                                                    setFieldValue("Description", text);
                                                                    return text;
                                                                }} />
                                                        </Field>
                                                    </div>
                                                </div>
                                            </Form>)
                                        }
                                    }
                                </Formik >
                                :
                                <Spinner size='medium' label={strings.Loading + "..."} />
                            }
                        </DialogContent>
                        <DialogActions>
                            <DialogTrigger action='close'>
                                <Button appearance='secondary' onClick={() => {
                                    this.setState({ keepTheFormOpen: false });
                                    this.props.openForm();
                                }}>
                                    {strings.Cancel}
                                </Button>
                            </DialogTrigger>
                            {
                                !locked && !formDataLoading &&
                                <>
                                    {
                                        !submitting &&
                                        <Button appearance='primary' form='convo' type='submit'> {strings.Save} </Button>
                                    }
                                    {
                                        submitting &&
                                        <Spinner size='extra-tiny' label={strings.Saving + "..."} />
                                    }
                                </>
                            }
                        </DialogActions>
                    </DialogBody>
                </DialogSurface>
            </Dialog>
        );
    }

}