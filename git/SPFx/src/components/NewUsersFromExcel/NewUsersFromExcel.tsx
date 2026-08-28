import * as React from 'react';
import {INewUsersErrors, INewUsersFromExcelprops, INewUsersFromExcelState } from './INewUsersFromExcel';
import styles from './NewUsersFromExcel.module.scss';
import * as XLSX from 'xlsx';
import { Button, createTableColumn, DataGrid, DataGridBody, DataGridCell, DataGridHeader, DataGridHeaderCell, DataGridRow, DialogActions, DialogContent, DialogTrigger, Divider, Dropdown, Label, Link,TableCellLayout, TableColumnDefinition, Option, Spinner, Textarea, Popover, PopoverTrigger, PopoverSurface} from '@fluentui/react-components';
import { Administration, BodyUserInfo, excelColumsMap, NewUserFromExcel } from '../../webparts/administrationApp/models/AdministrationAppModels';
import strings from 'NewUserFormStrings';
import { getBodyIdFormUrl } from '../../service/RoleService';
import { ErrorCircleRegular } from '@fluentui/react-icons';

class NewUsersFromExcel extends React.Component<INewUsersFromExcelprops, INewUsersFromExcelState> {
  private columnsUser: TableColumnDefinition<NewUserFromExcel>[] =
  this.props.isMobile ? 
  [ 
    createTableColumn<NewUserFromExcel>({
      columnId: "Nombre",
      renderHeaderCell: () => {
        return strings.Name;
      },
      renderCell: (item) => {
        return (
          <TableCellLayout className={styles.tableCell}>
            {item.Nombre}
          </TableCellLayout>
        );
      },
    }),createTableColumn<NewUserFromExcel>({
      columnId: "Apellidos",
      renderHeaderCell: () => {
        return strings.SurName;
      },
      renderCell: (item) => {
        return (
          <TableCellLayout className={styles.tableCell}>
            {item.Apellidos}
          </TableCellLayout>
        );
      },
    }),createTableColumn<NewUserFromExcel>({
      columnId: "Email",
      renderHeaderCell: () => {
        return strings.Email;
      },
      renderCell: (item) => {
        return (
          <TableCellLayout className={styles.tableCell}>
            {item.Email}
          </TableCellLayout>
        );
      },
    }),createTableColumn<NewUserFromExcel>({
      columnId: "Cargo",
      renderHeaderCell: () => {
        return strings.Job;
      },
      renderCell: (item) => {
        return (
          <TableCellLayout className={styles.tableCell}>
            {item.Cargo}
          </TableCellLayout>
        );
      },
    }),createTableColumn<NewUserFromExcel>({
      columnId: "ComunidadAutonoma",
      renderHeaderCell: () => {
        return strings.AutonomousCommunity;
      },
      renderCell: (item) => {
        return (
          <TableCellLayout className={styles.tableCell}>
            {item.ComunidadAutonoma}
          </TableCellLayout>
        );
      },
    }),createTableColumn<NewUserFromExcel>({
      columnId: "Rol",
      renderHeaderCell: () => {
        return strings.RoleLabel;
      },
      renderCell: (item) => {
        return (
          <TableCellLayout className={styles.tableCell}>
            {item.Rol}
          </TableCellLayout>
        );
      },
    }),
  ]:
  [
    createTableColumn<NewUserFromExcel>({
      columnId: "Nombre",
      renderHeaderCell: () => {
        return strings.Name;
      },
      renderCell: (item) => {
        return (
          <TableCellLayout>
            {item.Nombre}
          </TableCellLayout>
        );
      },
    }),createTableColumn<NewUserFromExcel>({
      columnId: "Apellidos",
      renderHeaderCell: () => {
        return strings.SurName;
      },
      renderCell: (item) => {
        return (
          <TableCellLayout>
            {item.Apellidos}
          </TableCellLayout>
        );
      },
    }),createTableColumn<NewUserFromExcel>({
      columnId: "Email",
      renderHeaderCell: () => {
        return strings.Email;
      },
      renderCell: (item) => {
        return (
          <TableCellLayout>
            {item.Email}
          </TableCellLayout>
        );
      },
    }),createTableColumn<NewUserFromExcel>({
    columnId: "CorreoElectronicoSecundario",
    renderHeaderCell: () => {
      return strings.SecondaryMail;
    },
    renderCell: (item) => {
      return (
        <TableCellLayout>
          {item.CorreoElectronicoSecundario}
        </TableCellLayout>
      );
    },
    }),createTableColumn<NewUserFromExcel>({
      columnId: "TelefonoEmtestsa",
      renderHeaderCell: () => {
        return strings.BusinessPhoneLabel;
      },
      renderCell: (item) => {
        return (
          <TableCellLayout>
            {item.TelefonoEmtestsa}
          </TableCellLayout>
        );
      },
    }),createTableColumn<NewUserFromExcel>({
      columnId: "TelefonoMovil",
      renderHeaderCell: () => {
        return strings.PhoneNumber;
      },
      renderCell: (item) => {
        return (
          <TableCellLayout>
            {item.TelefonoMovil}
          </TableCellLayout>
        );
      },
    }),createTableColumn<NewUserFromExcel>({
      columnId: "ComunidadAutonoma",
      renderHeaderCell: () => {
        return strings.AutonomousCommunity;
      },
      renderCell: (item) => {
        return (
          <TableCellLayout>
            {item.ComunidadAutonoma}
          </TableCellLayout>
        );
      },
    }),createTableColumn<NewUserFromExcel>({
      columnId: "Division_GobiernoAutonomico",
      renderHeaderCell: () => {
        return strings.CompanyNameLabel;
      },
      renderCell: (item) => {
        return (
          <TableCellLayout>
            {item.Division_GobiernoAutonomico}
          </TableCellLayout>
        );
      },
    }),createTableColumn<NewUserFromExcel>({
      columnId: "Unidad_Consejeria",
      renderHeaderCell: () => {
        return strings.DepartmentLabel;
      },
      renderCell: (item) => {
        return (
          <TableCellLayout>
            {item.Unidad_Consejeria}
          </TableCellLayout>
        );
      },
    }),createTableColumn<NewUserFromExcel>({
      columnId: "Cargo",
      renderHeaderCell: () => {
        return strings.Job;
      },
      renderCell: (item) => {
        return (
          <TableCellLayout>
            {item.Cargo}
          </TableCellLayout>
        );
      },
    }),createTableColumn<NewUserFromExcel>({
      columnId: "TipoEmpleado",
      renderHeaderCell: () => {
        return strings.EmployeeTypeLabel;
      },
      renderCell: (item) => {
        return (
          <TableCellLayout>
            {item.TipoEmpleado}
          </TableCellLayout>
        );
      },
    }),createTableColumn<NewUserFromExcel>({
      columnId: "Rol",
      renderHeaderCell: () => {
        return  strings.RoleLabel;
      },
      renderCell: (item) => {
        return (
          <TableCellLayout>
            {item.Rol}
          </TableCellLayout>
        );
      },
    }),createTableColumn<NewUserFromExcel>({
      columnId: "DestinatarioInvitacionCC",
      renderHeaderCell: () => {
        return strings.InvitationCCO;
      },
      renderCell: (item) => {
        return (
          <TableCellLayout>
            {item.DestinatarioInvitacionCC}
          </TableCellLayout>
        );
      },
    }),
  ];

  private columnsprocess: TableColumnDefinition<NewUserFromExcel>[] =
  this.props.isMobile ? 
  [ 
    createTableColumn<NewUserFromExcel>({
      columnId: "Nombre",
      renderHeaderCell: () => {
        return strings.Name;
      },
      renderCell: (item) => {
        return (
          <TableCellLayout className={styles.tableCell}>
            {item.Nombre}
          </TableCellLayout>
        );
      },
    }),createTableColumn<NewUserFromExcel>({
      columnId: "Apellidos",
      renderHeaderCell: () => {
        return strings.SurName;
      },
      renderCell: (item) => {
        return (
          <TableCellLayout className={styles.tableCell}>
            {item.Apellidos}
          </TableCellLayout>
        );
      },
    }),createTableColumn<NewUserFromExcel>({
      columnId: "Status",
      renderHeaderCell: () => {
        return strings.Status;
      },
      renderCell: (item) => {
        const currentStatus = this.state.correctRequest && this.state.correctRequest.includes(item.Email) ? strings.StatusOk : this.state.failRequest && this.state.failRequest.includes(item.Email) ? strings.StatusError: strings.StatusPending;

        return (
          <TableCellLayout className={styles.tableCell} style={{color: currentStatus === strings.StatusOk ? 'green': currentStatus === strings.StatusError ? '#bc2f32' : 'inherit' }}>
            {currentStatus}
          </TableCellLayout>
        );
      },
    })
  ]:
  [
    createTableColumn<NewUserFromExcel>({
      columnId: "Nombre",
      renderHeaderCell: () => {
        return strings.Name;
      },
      renderCell: (item) => {
        return (
          <TableCellLayout>
            {item.Nombre}
          </TableCellLayout>
        );
      },
    }),createTableColumn<NewUserFromExcel>({
      columnId: "Apellidos",
      renderHeaderCell: () => {
        return strings.SurName;
      },
      renderCell: (item) => {
        return (
          <TableCellLayout>
            {item.Apellidos}
          </TableCellLayout>
        );
      },
    }),createTableColumn<NewUserFromExcel>({
      columnId: "Email",
      renderHeaderCell: () => {
        return strings.Email;
      },
      renderCell: (item) => {
        return (
          <TableCellLayout>
            {item.Email}
          </TableCellLayout>
        );
      },
    }),createTableColumn<NewUserFromExcel>({
      columnId: "Status",
      renderHeaderCell: () => {
        return strings.Status;
      },
      renderCell: (item) => {
        const currentStatus = this.state.correctRequest && this.state.correctRequest.includes(item.Email) ? strings.StatusOk : this.state.failRequest && this.state.failRequest.includes(item.Email) ? strings.StatusError: strings.StatusPending;

        return (
          <TableCellLayout className={styles.tableCell} style={{color: currentStatus === strings.StatusOk ? 'green': currentStatus === strings.StatusError ? '#bc2f32' : 'inherit' }}>
            {currentStatus}
          </TableCellLayout>
        );
      },
    })
  ];

  constructor(props: INewUsersFromExcelprops | Readonly<INewUsersFromExcelprops>) {
    super(props);
    this.state = {
      newUsers:[],
      showConfirmPanel: false,
      correctRequest: [],
      failRequest: [],
      isprocessing: false,
      loading: false,
      showErrorUsersNotprocessed: false,
      invitationMessage: Administration.invitationMessage,
      errors: []
    };
  }

  componentDidMount(): void {
    this.setState({invitationMessage:this.props.invitationMsg ? this.props.invitationMsg : Administration.invitationMessage});
  }

  handleChange(selectorFiles: FileList)
  { 
    if(selectorFiles && selectorFiles.length > 0 ){
      const file = selectorFiles[0];
      const reader = new FileReader();
      reader.onload = (event) => {
        if(event && event.target && event.target.result){
          console.log(event);
          const workbook = XLSX.read(event.target.result, {type: 'binary'});
          const sheetName = workbook.SheetNames[0];
          const sheet = workbook.Sheets[sheetName];
          const jsonData = XLSX.utils.sheet_to_json(sheet);
          console.log(jsonData);
          const jsonDataClean = this.limpiarPersonas(jsonData);
          const users: NewUserFromExcel[] =
            this.props.testSelectedOrgan ? 
              jsonDataClean
              .filter((item: any) => item[excelColumsMap.Nombre] && item[excelColumsMap.Apellidos] && item[excelColumsMap.Email] && item[excelColumsMap.Email].toString().match(/^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/) && item[excelColumsMap.Rol])
              .map((item: any) => ({
                Nombre: item[excelColumsMap.Nombre] || "",
                Apellidos: item[excelColumsMap.Apellidos] || "",
                Email: item[excelColumsMap.Email] && item[excelColumsMap.Email].toString().match(/^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/) ? item[excelColumsMap.Email].toLowerCase() : "",
                Cargo: item[excelColumsMap.Cargo] || "",
                ComunidadAutonoma: item[excelColumsMap.ComunidadAutonoma] || "",
                Rol: item[excelColumsMap.Rol] || undefined,
                Division_GobiernoAutonomico: item[excelColumsMap.Division_GobiernoAutonomico] || "",
                Unidad_Consejeria: item[excelColumsMap.Unidad_Consejeria] || "",
                TelefonoEmtestsa: item[excelColumsMap.TelefonoEmtestsa] && item[excelColumsMap.TelefonoEmtestsa].toString().match(/^[0-9]+$/) ? item[excelColumsMap.TelefonoEmtestsa] : "",
                TelefonoMovil: item[excelColumsMap.TelefonoMovil] && item[excelColumsMap.TelefonoMovil].toString().match(/^[0-9]+$/) ? item[excelColumsMap.TelefonoMovil] :  "",
                CorreoElectronicoSecundario: item[excelColumsMap.CorreoElectronicoSecundario] && item[excelColumsMap.CorreoElectronicoSecundario].toString().match(/^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/) ? item[excelColumsMap.CorreoElectronicoSecundario]: "",
                TipoEmpleado: item[excelColumsMap.TipoEmpleado] || "", 
                DestinatarioInvitacionCC: item[excelColumsMap.DestinatarioInvitacionCC] && item[excelColumsMap.DestinatarioInvitacionCC].toString().match(/^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/) ? item[excelColumsMap.DestinatarioInvitacionCC] : ""
              }))
            :
            jsonDataClean
            .filter((item: any) => item[excelColumsMap.Nombre]  && item[excelColumsMap.Apellidos] && item[excelColumsMap.Email] && item[excelColumsMap.Email].toString().match(/^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/))
            .map((item: any) => ({
              Nombre: item[excelColumsMap.Nombre] || "",
              Apellidos: item[excelColumsMap.Apellidos] || "",
              Email: item[excelColumsMap.Email] && item[excelColumsMap.Email].toString().match(/^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/) ? item[excelColumsMap.Email].toLowerCase() : "",
              Cargo: item[excelColumsMap.Cargo] || "",
              ComunidadAutonoma: item[excelColumsMap.ComunidadAutonoma] || "",
              Rol: item[excelColumsMap.Rol] || undefined,
              Division_GobiernoAutonomico: item[excelColumsMap.Division_GobiernoAutonomico] || "",
              Unidad_Consejeria: item[excelColumsMap.Unidad_Consejeria] || "",
              TelefonoEmtestsa: item[excelColumsMap.TelefonoEmtestsa] && item[excelColumsMap.TelefonoEmtestsa].toString().match(/^[0-9]+$/) ? item[excelColumsMap.TelefonoEmtestsa] : "",
              TelefonoMovil: item[excelColumsMap.TelefonoMovil] && item[excelColumsMap.TelefonoMovil].toString().match(/^[0-9]+$/) ? item[excelColumsMap.TelefonoMovil] :  "",
              CorreoElectronicoSecundario: item[excelColumsMap.CorreoElectronicoSecundario] && item[excelColumsMap.CorreoElectronicoSecundario].toString().match(/^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/) ? item[excelColumsMap.CorreoElectronicoSecundario]: "",
              TipoEmpleado: item[excelColumsMap.TipoEmpleado] || "", 
              DestinatarioInvitacionCC: item[excelColumsMap.DestinatarioInvitacionCC] && item[excelColumsMap.DestinatarioInvitacionCC].toString().match(/^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/) ? item[excelColumsMap.DestinatarioInvitacionCC] : ""
            }));
          const errors = this.usersWithError(jsonDataClean);
          let hasErrorprocessing= false;
          if(this.props.testSelectedOrgan){
            hasErrorprocessing =  jsonDataClean.some((item: any) => !item[excelColumsMap.Nombre] || !item[excelColumsMap.Apellidos] || !item[excelColumsMap.Email] || !item[excelColumsMap.Rol] || !item[excelColumsMap.Email]?.toString().match(/^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/))
          }else{
            hasErrorprocessing =  jsonDataClean.some((item: any) => !item[excelColumsMap.Nombre] || !item[excelColumsMap.Apellidos] || !item[excelColumsMap.Email] || !item[excelColumsMap.Email]?.toString().match(/^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/))
          }
          this.setState({newUsers: users.filter(us => !errors.some(er => er.Email&&er.Email.toLocaleLowerCase() == us.Email)), showErrorUsersNotprocessed: (hasErrorprocessing ||errors.length > 0), errors})
        }
      };
      reader.readAsArrayBuffer(file);
    }
  }

  
  private limpiarPersonas = (personas: any[]): any[] => {
    return personas.filter(persona => {
      const {[excelColumsMap.TipoEmpleado]: tipoEmpleado, ...resto } = persona;

      const todosVacios = Object.values(resto).every(
        valor => valor === undefined || valor === null || valor === ""
      );

      return !(todosVacios && tipoEmpleado === "");
    });
  };

  private usersWithError = (personas: any[]): INewUsersErrors[] => {
      
    return personas.reduce((errores: INewUsersErrors[], usuario: any) => {
        const error: INewUsersErrors = {};

        const nombre = usuario[excelColumsMap.Nombre];
        const apellidos = usuario[excelColumsMap.Apellidos];
        const cargo = usuario[excelColumsMap.Cargo];
        const comunidad = usuario[excelColumsMap.ComunidadAutonoma];
        const Division = usuario[excelColumsMap.Division_GobiernoAutonomico];
        const consejeria = usuario[excelColumsMap.Unidad_Consejeria];
        const tipoEmpleado = usuario[excelColumsMap.TipoEmpleado];
        const mail = usuario[excelColumsMap.Email];
        const mailSecundario = usuario[excelColumsMap.CorreoElectronicoSecundario];
        const telefonoEmtestsa = usuario[excelColumsMap.TelefonoEmtestsa];
        const telefonoMovil = usuario[excelColumsMap.TelefonoMovil];
        const rol = usuario[excelColumsMap.Rol];
        if (!nombre || nombre.trim() === "") {
          error.Nombre = strings.Name + ": "+strings.FieldRequired;
        }

        if (!apellidos || apellidos.trim() === "") {
          error.Apellidos = strings.SurName + ": "+strings.FieldRequired;
        }

        if (cargo && cargo.toString().length > 128){
          error.Cargo = strings.Job + ": "+strings.MaxLong128;
        }

        if (comunidad && comunidad.toString().length > 64) {
          error.ComunidadAutonoma = strings.AutonomousCommunity + ": "+strings.MaxLong64;
        }

        if (Division && Division.toString().length > 64) {
          error.Division_GobiernoAutonomico = strings.CompanyNameLabel + ": "+strings.MaxLong64;
        }
        
        if (consejeria && consejeria.toString().length > 64) {
          error.Unidad_Consejeria = strings.DepartmentLabel + ": "+strings.MaxLong64;
        }

        if (tipoEmpleado && tipoEmpleado.toString().length > 64) {
          error.TipoEmpleado = strings.EmployeeTypeLabel + ": "+strings.MaxLong64;
        }

        if(mailSecundario && !mailSecundario.toString().match(/^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/)){
          error.CorreoElectronicoSecundario = strings.SecondaryMail + ": "+strings.MailError;
        }

        if(telefonoEmtestsa && !telefonoEmtestsa.toString().match(/^[0-9]+$/)){
          error.TelefonoEmtestsa = strings.BusinessPhoneLabel + ": "+strings.PhoneError;
        }

        if(telefonoMovil && !telefonoMovil.toString().match(/^[0-9]+$/)){
          error.TelefonoMovil = strings.PhoneNumber + ": "+strings.PhoneError;
        }

        if(this.props.testSelectedOrgan && !rol ){
          error.Rol = strings.RoleLabel + ": "+strings.FieldRequired;
        }

        // Solo agregamos el error si hay al menos una prodpiedad con mensaje
        if (Object.keys(error).length > 0) {
          error.Email = mail;
          errores.push(error);
        }

        return errores;
      }, []);

  };

  private onMaprodleToBack(roleTaxId: string): string {
    if (roleTaxId === "Convocante") {
      return "schedulers";
    } else if (roleTaxId === "Gestor del convocante") {
      return "gestorschedulers";
    } else if (roleTaxId === "Miembro") {
      return "members";
    } else if (roleTaxId === "Asistente del miembro") {
      return "asistentemembers";
    } else if (roleTaxId === "Invitado") {
      return "guests";
    } else {
      return roleTaxId;
    }
  }

  private onprocessUsers (users:BodyUserInfo[]): void {
    const allpromises:promise<void>[] = [];
    const bodyId: string = getBodyIdFormUrl(this.props.context.pageContext.site.serverRelativeUrl);
    users.forEach(user => {
      const promise = this.props.bkService.createUser(user, bodyId || "Contoso").then(()=> {
        if(user && user.UserInfo && user.UserInfo.PrincipalMail){
          this.setState((testvState) => ({
            correctRequest:[...testvState.correctRequest, user.UserInfo.PrincipalMail || ""]
          }))
        }
      }).catch((ex)=> {
        console.log(ex);
        if(user && user.UserInfo && user.UserInfo.PrincipalMail){
          this.setState((testvState) => ({
            failRequest:[...testvState.failRequest, user.UserInfo.PrincipalMail || ""]
          }))
        }
      })
      allpromises.push(promise);
    });
    promise.all(allpromises).then(()=> {
      console.log("promises completed");
    }).catch((ex)=> {
      console.log(ex)
    }).finally(()=> this.setState({loading: false}))
  }

  private onClickClose(): void{
    const { showConfirmPanel } = this.state;
    if(!showConfirmPanel){
      this.props.onClosePanel();
      //this.setState({formErrors:{}, userInfo:undefined, showConfirmPanel: false, isLoadingSave: false})
    }else{
       this.setState({showConfirmPanel: false});
    }
  }
  
  private onClickSave(): void{
    const { newUsers, showConfirmPanel, selectedOrgan ,isprocessing, invitationMessage} = this.state;
    if(!showConfirmPanel){
      /*const hasError = this.onCheckErrors();*/
      const hasError = false;
      if(newUsers && !hasError){
        this.setState({showConfirmPanel:true})
      }
      
    }else{
      if(!isprocessing && newUsers && newUsers.length > 0){
        const processUsers:BodyUserInfo[] = newUsers.map(user => ({
          UserInfo:{
            FirstName: user.Nombre,
            LastName: user.Apellidos,
            JobTitle: user.Cargo || "",
            Office: user.ComunidadAutonoma || "",
            PrincipalMail: user.Email.toLowerCase(),
            BusinessPhone: user.TelefonoEmtestsa,
            CellPhone: user.TelefonoMovil,
            CompanyName: user.Division_GobiernoAutonomico,
            Department: user.Unidad_Consejeria,
            Email: user.CorreoElectronicoSecundario,
            EmployeeType: user.TipoEmpleado,
          },
          UserBodyRole:{
            UserUpn:"",
            BodyId: ((selectedOrgan || this.props.testSelectedOrgan) && user.Rol) ? this.props.testSelectedOrgan ? this.props.testSelectedOrgan.key :selectedOrgan : "",
            Role: ((selectedOrgan || this.props.testSelectedOrgan)&& user.Rol) ? this.onMaprodleToBack(user.Rol) : "",
          },
          UserInvitation:{
            InvitationMessage: invitationMessage,
            InvitationCCRecipient: user.DestinatarioInvitacionCC
          }
        }));

        console.log(processUsers);
        this.setState({isprocessing: true, loading: true});
        //if(this.state.isprocessing)
        this.onprocessUsers(processUsers);
        /*this.setState({isLoadingSave: true});
        let userToCreate:BodyUserInfo = {...userInfo}; 
        if(!userInfo.UserBodyRole || !userInfo?.UserBodyRole?.Role || !userInfo?.UserBodyRole?.BodyId){
          userToCreate.UserBodyRole = {
            UserUpn:"",
            Role:"",
            BodyId:""
          };
        }else if(userToCreate && userToCreate.UserBodyRole && userToCreate.UserBodyRole.BodyId && userToCreate.UserBodyRole.Role){
          userToCreate.UserBodyRole.UserUpn = "";
        }
        this.props.onCreateUser(newUsers);
        */
        //this.setState({formErrors:{}, userInfo:undefined, showConfirmPanel: false, isLoadingSave: false})
      }
      if(isprocessing){
        this.props.onRealoadComponent();
      }
    }
  }

  public render(): JSX.Element {
    const { newUsers, selectedOrgan, showConfirmPanel,isprocessing,loading, showErrorUsersNotprocessed, invitationMessage} = this.state;
    const { isMobile, organFilterOptions, testSelectedOrgan } = this.props;
    return (
      <>
      <DialogContent>
        {showConfirmPanel ? 
          <div className={`${styles.newUsersFromExcel}`}>
            {!isprocessing ? 
              <>{strings.AreYouSure}
              </>
            :
              <div className={styles.resultExcelZone}>
                {loading &&
                  <Spinner/>
                }
                <DataGrid
                  items={newUsers}
                  columns={this.columnsprocess}
                  className={styles.dataGridContainer}
                  noNativeElements={isMobile}
                  sortable
                >
                  <DataGridHeader className={styles.dataGridHeadersContainer}>
                    <DataGridRow
                    >
                      {({ renderHeaderCell }) => (
                        <DataGridHeaderCell className={styles.dataGridHeaders}>{renderHeaderCell()}</DataGridHeaderCell>
                      )}
                    </DataGridRow>
                  </DataGridHeader>
                  <DataGridBody<NewUserFromExcel> className={styles.dataGridBody}>
                    {({ item, rowId }) =>
                    (
                      <DataGridRow<NewUserFromExcel>
                        key={rowId}
                        className={styles.dataRowItems}                >
                        {({ renderCell, columnId }) => (
                          <DataGridCell focusMode={"none"} className={isMobile && columnId === "openItem" ? styles.manageItem : ''}>
                            {renderCell(item)}
                          </DataGridCell>
                        )}
                      </DataGridRow>
                    )}
                  </DataGridBody>
                </DataGrid>
              </div>
            }
          </div>
        :
          <div className={`${styles.newUsersFromExcel}`}>
            <div className={styles.uploadFileZone}>
              <div className={styles.uploadRow}>
                <div style={{overflow:"hidden"}}>
                  <input 
                    type="file" 
                    onChange={ (e) => e?.target?.files && this.handleChange(e.target.files) }
                    accept="application/vnd.openxmlformats-officedocument.stestadsheetml.sheet, application/vnd.ms-excel"
                  />
                </div>
                <div>
                  <Link 
                    //href={"/sites/Contoso/_layouts/download.aspx?SourceUrl="+window.location.origin+"/sites/Contoso/Documents%20compartidos/CargaUsuarios_Template.xlsx"} 
                    target='_blank'
                    onClick={(e) => {
                      e.stopprodpagation(); 
                      e.preventDefault()
                      window.open("/sites/Contoso/_layouts/download.aspx?SourceUrl="+window.location.origin+"/sites/Contoso/Documents%20compartidos/CargaUsuarios_Template.xlsx",'_blank')
                    }}
                  >{strings.DownloadTemplate}
                  </Link>
                </div>
              </div>
              <div className={styles.tooltipComment}>
                {testSelectedOrgan ? strings.MandatoryFieldOrgans : strings.MandatoryFieldsWarning}
                <br/>
                {strings.MandatoryLengthFields}
              </div>  
            </div>
            <Divider></Divider>
            <div className={styles.gridContainerSingle}>
              {/* Órgano */}
              <div className={styles.propertyRow}>
                <Label className={styles.labelproperty}>
                  {strings.DepartmentName}
                </Label>
                <Dropdown
                  className={styles.dropDownValue}
                  value={selectedOrgan ? organFilterOptions.find(or => or.key === selectedOrgan)?.text : testSelectedOrgan ? testSelectedOrgan.text : undefined}
                  onOptionSelect={(ev, data) => this.setState({selectedOrgan:data.optionValue ? data.optionValue : ""})}
                  clearable={testSelectedOrgan ? false: true}
                >
                  {
                    organFilterOptions.map(option =>
                      <Option
                        key={option.key}
                        value={option.key}
                      >
                        {option.text}
                      </Option>
                    )
                  }
                </Dropdown>
              </div>
              <div className={styles.tooltipComment}>
                {testSelectedOrgan ? strings.OrganRoleMandatoryTooltip : strings.AsignOrganRoleTooltip}
              </div>
            </div>
            <Divider></Divider>
            <div className={styles.resultExcelZone}>
              {showErrorUsersNotprocessed &&
                <div className={styles.usersNotprocessed}>
                  {strings.UsersNotprocessed}
                  {this.state.errors && this.state.errors.length > 0 && 
                    <Popover>
                      <PopoverTrigger disableButtonEnhancement>
                        <Button icon={<ErrorCircleRegular />} className={styles.usersNotprocessedButton}></Button>
                      </PopoverTrigger>

                      <PopoverSurface tabIndex={-1} style={{height:"250px", overflowY:"auto"}}>
                        <div>
                          <ul>
                          {this.state.errors.map( er => 
                            <li><strong>{er.Email}</strong>
                              <ul>
                              {Object.entries(er)
                              .filter(([key, value]) => key !== "Email" && value) // excluye "Email" y valores vacíos
                              .map(([key, value], i) => (
                                <li key={i}>
                                  {value}
                                </li>
                              ))}
                              </ul>
                            </li>
                          )
                          }
                          </ul>
                        </div>
                      </PopoverSurface>
                    </Popover>
                  }
                </div>
              }
              {newUsers && newUsers.length > 0 ? 
                <DataGrid
                  items={newUsers}
                  columns={this.columnsUser}
                  className={styles.dataGridContainer}
                  noNativeElements={isMobile}
                  sortable
                >
                  <DataGridHeader className={styles.dataGridHeadersContainer}>
                    <DataGridRow
                    >
                      {({ renderHeaderCell }) => (
                        <DataGridHeaderCell className={styles.dataGridHeaders}>{renderHeaderCell()}</DataGridHeaderCell>
                      )}
                    </DataGridRow>
                  </DataGridHeader>
                  <DataGridBody<NewUserFromExcel> className={styles.dataGridBody}>
                    {({ item, rowId }) =>
                    (
                      <DataGridRow<NewUserFromExcel>
                        key={rowId}
                        className={styles.dataRowItems}                >
                        {({ renderCell, columnId }) => (
                          <DataGridCell focusMode={"none"} className={isMobile && columnId === "openItem" ? styles.manageItem : ''}>
                            {renderCell(item)}
                          </DataGridCell>
                        )}
                      </DataGridRow>
                    )}
                  </DataGridBody>
                </DataGrid>
                :
                <div className={styles.noItems}>
                  {strings.NoUsersLoaded}
                </div>
              }
            </div>
            <Divider />
             <div className={styles.gridContainerSingle}>
                {/* invitation message */}
                <div className={styles.propertyRow}>
                  <Label className={styles.labelproperty} required>
                    {strings.InvitationMessage}
                  </Label>
                    <Textarea 
                      className={!invitationMessage ? `${styles.dropDownValue} ${styles.errorField} ${styles.textArea}` : `${styles.dropDownValue} ${styles.textArea}`}
                      value={invitationMessage}
                      onChange={(ev, data) => {this.setState({invitationMessage: data.value})}}
                      resize='both'
                    />
                </div>
              </div>
          </div>
        }
      </DialogContent>
      <DialogActions>
        <DialogTrigger disableButtonEnhancement>
          <Button
            appearance="secondary"
            onClick={() => this.onClickClose()}
            disabled={loading}
            style={{visibility: isprocessing ? 'hidden': 'visible'}}
          >
            {strings.Discard}
          </Button>
        </DialogTrigger>
          <Button
            appearance="primary"
            onClick={() => this.onClickSave()}
            disabled={!newUsers || newUsers.length < 1 || loading || !invitationMessage}
            //disabled={isLoadingSave||showError}
          >
            {!isprocessing ? strings.Save : strings.Close}
          </Button>
      </DialogActions >
    </>
    );
  }
}

export default NewUsersFromExcel;
