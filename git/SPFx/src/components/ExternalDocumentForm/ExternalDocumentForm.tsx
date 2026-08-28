import * as React from 'react';
import { IEXTERNALDoc, IEXTERNALDocumentFormprops, IEXTERNALDocumentFormState } from './IEXTERNALDocumentForm';
import styles from './EXTERNALDocumentForm.module.scss';
import { Button, Checkbox, Combobox, Dialog, DialogActions, DialogBody, DialogSurface, DialogTitle, Input, Label, Option, Spinner } from '@fluentui/react-components';
import { IFolderInfo } from '@pnp/sp/folders/types';
import strings from 'CertificationAppWebPartStrings';

class EXTERNALDocumentForm extends React.Component<IEXTERNALDocumentFormprops, IEXTERNALDocumentFormState> {
  constructor(props: IEXTERNALDocumentFormprops | Readonly<IEXTERNALDocumentFormprops>) {
    super(props);
    this.state = {
      document:{
        TipoDocumentLookup: "",
        TipoDepartmentLookup: "",
        Department: "",
        FechaMeeting: "",
      },
      isLoading: false,
      filteredOrgans:[],
      folders:[],
      filterOrganTxt: "",
      errors: {},
      saved: false,
      hasError: false
    };
  }
    
  componentDidMount(): void {
    this.setState({filteredOrgans:this.props.organs})
  }

  handleChange = (field: keyof IEXTERNALDoc, value: string|boolean) => {
    this.setState(
      (testv) => ({
        document: { ...testv.document, [field]: value },
      }),
    );
  };

  
  handleSubmit = async () => {
    // 1) Validación
    const errors = this.validate(this.state.document);
    if (Object.keys(errors).length > 0) {
      this.setState({ errors });
      return;
    }

    // 2) Encender loading
    this.setState({errors:{}, isLoading: true, saved: false, hasError: false });

    try {
      // 3) Resolver carpeta
      let folder: IFolderInfo | undefined;
      if (this.state.document.NuevaCarpeta) {
        const createFolder = await this.props.spService.ensureFolder(
          `${this.state.Departmentseleccionado?.FileRef}/${this.state.document.NombreNuevaCarpeta}`
        );
        folder = await createFolder.folder();
      } else {
        folder = this.state.selectedFolder;
      }

      // 4) Subir archivo (IMPORTANTE: await)
      if (folder && this.state.file) {
        const uploadedFile = await this.props.spService.uploadFile(
          folder.ServerRelativeUrl,
          this.state.file
        );

        const item = await uploadedFile.file.getItem();
        await item.update({
          TipoDocumentLookupId: Number(this.state.document.TipoDocumentLookup),
          FechaMeeting: this.state.document.FechaMeeting
            ? new Date(this.state.document.FechaMeeting)
            : null,
        });

        // 5) Estado final OK
        this.setState({ isLoading: false, saved: true, hasError: false, file: undefined });
      } else {
        // Si por alguna razón no hay folder o file tras validar
        this.setState({ isLoading: false, saved: false, hasError: true });
      }
    } catch (error) {
      console.error("Error al subir archivo:", error);
      // 6) Estado error
      this.setState({ isLoading: false, saved: false, hasError: true });
    }
  };


  handleFile = (filelist: FileList) => {
    this.setState({file:filelist&& filelist.length>0 ? filelist[0]: undefined, saved: false, hasError: false, isLoading: false});
  };

  updateOrganValue = async (id:string) => {
    const selectedOrgan = this.props.organs.find(or => or.ID?.toString() === id.toString());
    if(selectedOrgan){
      this.setState((testvState) => ({
      document: {
        ...testvState.document,
        "Department": selectedOrgan.Department,
        "TipoDepartmentLookup": selectedOrgan.TipoDepartmentEXTERNAL,
        "CarpetaReunion": undefined
      },
      Departmentseleccionado: selectedOrgan,
      filteredOrgans:this.props.organs,
      filterOrganTxt: selectedOrgan.NombreDepartment,
      selectedFolder: undefined
    }));
    if(selectedOrgan.FileRef){
      const folders:IFolderInfo[] = await this.props.spService.getFoldersFromPath(selectedOrgan.FileRef);
      this.setState({folders});
    }
    }
  }
  
  onFilterOrgan = (text:string)=>{
    if(text){
      const newOrgans = this.props.organs.filter(or => or.NombreDepartment.toLowerCase().includes(text.toLowerCase()));
      this.setState({filteredOrgans:newOrgans});

    }else{
      this.setState({filteredOrgans:this.props.organs})
    }
  }

  
  validate = (doc: IEXTERNALDoc): Record<string, string> => {
    const errors: Record<string, string> = {};
    if (!this.state.file) errors.File = strings.FileRequired;
    if (!doc.TipoDocumentLookup) errors.TipoDocumentLookup = strings.TipoDocumentRequired;
    if (!doc.TipoDepartmentLookup) errors.TipoDepartmentLookup = strings.TipoDepartmentRequired;
    if (!doc.Department) errors.Department = strings.DepartmentRequired;
    //if (!doc.FechaMeeting) errors.FechaMeeting = "Fecha de Meeting es obligatoria.";

    if (!doc.CarpetaReunion) {
      if (!doc.NuevaCarpeta) {
        errors.CarpetaReunion = strings.SelectFolderRequired;
      } else if (!doc.NombreNuevaCarpeta) {
        errors.NombreNuevaCarpeta = strings.NewFolderRequired;
      }else if (this.state.folders.some(f => f.Name.toLowerCase() === doc.NombreNuevaCarpeta?.toLowerCase())){
        errors.NombreNuevaCarpeta = strings.FolderDuplicated;
      }
    }else if(doc.NuevaCarpeta && !doc.NombreNuevaCarpeta){
      errors.NombreNuevaCarpeta = strings.NewFolderRequired;
    }
    return errors;
  };


  public render(): JSX.Element {
    const {filteredOrgans, Departmentseleccionado, folders, filterOrganTxt, selectedFolder, errors, isLoading, saved, hasError} = this.state;
    const {organTypes, documentTypes} = this.props;
    return (
      <Dialog open={true}>
        <DialogSurface>
          <DialogTitle>{strings.UploadDocument}</DialogTitle>
          <DialogBody>
          <div className={styles.modalContent}>        
            {/* Input para subir archivo */}
            <div className={styles.propertyRow}>
              <Label className={styles.propertyLabel} htmlFor="fileUpload">{strings.FileLabel}</Label>
              <input 
                  type="file" 
                  onChange={ (e) => e?.target?.files && this.handleFile(e.target.files) }
                  disabled={isLoading}
                />
              {errors.File && <span className={styles.errorTxt}>{errors.File}</span>}
              {/* Dropdown Órgano */}
            </div>
            <div className={styles.propertyRow}>
              <Label className={styles.propertyLabel} required>Órgano</Label>
              <Combobox
                placeholder={strings.SelectOrganPlaceholder}
                selectedOptions={[Departmentseleccionado?.ID || ""]}
                onOptionSelect={(_, data) =>
                  this.updateOrganValue(data.optionValue|| "")
                }
                value={filterOrganTxt || ""}
                disabled={isLoading}
                onChange={(e) => {
                  const text = e.target.value;
                  this.setState({filterOrganTxt: text})
                  this.onFilterOrgan(text);
                }}
              >
                {filteredOrgans.map(org => 
                  <Option value={org.ID}>{org.NombreDepartment}</Option>
                )}
              </Combobox>
              {errors.Department && <span className={styles.errorTxt}>{errors.Department}</span>}
            </div>
            {/* Dropdown Tipo Órgano */}
            <div className={styles.propertyRow}>
              <Label className={styles.propertyLabel} required>{strings.OrganType}</Label>
                <Input
                value={this.state.document.TipoDepartmentLookup ? organTypes[this.state.document.TipoDepartmentLookup] :""}
                onChange={(e) => this.handleChange("TipoDepartmentLookup", e.target.value)}
                disabled
              />
              {errors.TipoDepartmentLookup && <span className={styles.errorTxt}>{errors.TipoDepartmentLookup}</span>}
            </div>
            <div className={styles.propertyRow}>
              {/* Dropdown Tipo Document */}
              <Label className={styles.propertyLabel} required>{strings.DocumentTypeLabel}</Label>
              <Combobox
                placeholder={strings.DocumentTypePlaceHolder}
                selectedOptions={[this.state.document.TipoDocumentLookup]}
                onOptionSelect={(_, data) =>
                  this.handleChange("TipoDocumentLookup", data.optionValue || "")
                }
                value={this.state?.document?.TipoDocumentLookup ? documentTypes[this.state.document.TipoDocumentLookup] : undefined}
                disabled={isLoading}
              >
                {Object.entries(documentTypes).map(([key, label]) => (
                    <Option key={key} value={key}>
                      {label}
                    </Option>
                  ))}
              </Combobox>
              {errors.TipoDocumentLookup && <span className={styles.errorTxt}>{errors.TipoDocumentLookup}</span>}
            </div>
            <div className={styles.propertyRow}>
              <Label className={styles.propertyLabel}>{strings.FechaMeetingLabel}</Label>
              <Input
                type="date"
                value={this.state.document.FechaMeeting}
                onChange={(e) => this.handleChange("FechaMeeting", e.target.value)}
                disabled={isLoading}
              />  
            </div>
            <div className={styles.propertyRow}>
              <Label className={styles.propertyLabel} required={!this.state.document.NuevaCarpeta}>{strings.CarpetaReunionLabel}</Label>
              <Combobox
                disabled={!Departmentseleccionado || this.state.document.NuevaCarpeta || isLoading}
                placeholder={strings.SelectFolderPlaceholder}
                selectedOptions={[this.state.document.CarpetaReunion || ""]}
                onOptionSelect={(_, data) =>
                  {
                  const folder = folders.find(f=> f.ServerRelativeUrl === data.optionValue);
                  this.setState({selectedFolder:folder});
                  this.handleChange("CarpetaReunion", data.optionValue || "")
                  }
                }
                value={selectedFolder?.Name || undefined}
              >
              {folders.map((folder)=> 
                <Option key={folder.Name} value={folder.ServerRelativeUrl}>
                  {folder.Name}
                </Option>
              )}
              </Combobox>
              {errors.CarpetaReunion && <span className={styles.errorTxt}>{errors.CarpetaReunion}</span>}
            </div>
            <div className={styles.propertyRow}>
              <Label>
                <Checkbox
                    checked={this.state.document.NuevaCarpeta || false}
                    onChange={(ev, data) => this.handleChange("NuevaCarpeta", ev.target.checked)}
                    disabled={!Departmentseleccionado || isLoading}
                />
                {strings.CreateNewFolder}
              </Label>
              {this.state.document.NuevaCarpeta && (
                <>
                  <Label className={styles.propertyLabel} required>{strings.NewFolderName}</Label>
                  <Input
                    value={this.state.document.NombreNuevaCarpeta || ""}
                    onChange={(e) =>
                      this.handleChange("NombreNuevaCarpeta", e.target.value)
                    }
                    disabled={isLoading}
                  />
                  {errors.NombreNuevaCarpeta && <span className={styles.errorTxt}>{errors.NombreNuevaCarpeta}</span>}
                </>
              )}
            </div>
          </div>   
        </DialogBody>
        <DialogActions style={{justifyContent:"space-between", justifySelf:"auto"}}>
          <div className={styles.bottomStatus}>
            {isLoading ? (
              <Spinner labelPosition="after" label={strings.SavingLoad} />
            ) : saved ? (
              <span>{strings.SavedOk}</span>
            ) : hasError ? (
              <span className={styles.errorTxt}>{strings.ErrorMsg}</span>
            ) : (
              <span />
            )}
          </div>
          <div>
            <Button appearance="secondary" onClick={this.props.onClose} disabled={isLoading}>
              {saved ? strings.Close : strings.Cancel}
            </Button>
            <Button appearance="primary" onClick={this.handleSubmit} disabled={isLoading || saved} style={{ marginLeft: "10px" }}>
              {strings.Save}
            </Button>
          </div>
        </DialogActions>
      </DialogSurface>
    </Dialog>
    );
  }
}

export default EXTERNALDocumentForm;
