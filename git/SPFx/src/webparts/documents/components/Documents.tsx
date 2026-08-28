import * as React from 'react';
import * as strings from 'DocumentsWebPartStrings';
import styles from './Documents.module.scss';
import { IDocumentsprops } from './IDocumentsprops';
import { IDocumentsState } from './IDocumentsState';
import { Spinner } from '@fluentui/react-components';
import { File, FileList, MgtTemplateprops } from '@microsoft/mgt-react/dist/es6/spfx';
import { ViewType } from '@microsoft/mgt-spfx';
import { DriveItem } from '@microsoft/microsoft-graph-types';


export const FileTemplate = (props: MgtTemplateprops) => {
  const file: DriveItem | undefined = props.dataContext ? props.dataContext.file : undefined;

  return (
    <File view={ViewType.oneline} fileDetails={file} className={styles.file}></File>
  );
}

export const Loading = (props: MgtTemplateprops) => {
  return (
    <div>
      <Spinner label={strings.Loading} labelPosition="after" size="small" />
    </div>
  );
}

export const NoData = (props: MgtTemplateprops) => {
  return (
    <span className={styles.noDocumentsAvailable}>{strings.NoDocumentsAvailable}</span>
  );
}

export default class RecentDocuments extends React.Component<IDocumentsprops, IDocumentsState> {

  constructor(props: IDocumentsprops) {
    super(props);
    this.state = {
      breadCrumbItems: [{ title: "Inicio", id: undefined }]
    };
  }

  public renderMoreLink = (): React.ReactElement => {
    return <a className={styles.moreLink} href={this.props.moreLink} title={strings.ViewMoreRelevantDocuments}> {strings.MoreLink} </a>;
  };

  public render(): React.ReactElement<IDocumentsprops> {
    const { title, moreLink, context } = this.props;
    const { currentDriveItemId, breadCrumbItems } = this.state;
    return (
      <section className={styles.documents}>
        <div className={styles.titleRow}>
          <div className={styles.wpTitle}>{title}</div>
          {
            moreLink && this.renderMoreLink()
          }
        </div>
        <div className={styles.mainContainer}>
          <ul className={styles.breadcrumb}>
            {
              breadCrumbItems?.map((item, index) => {
                return (
                  <li key={item.id}><a onClick={() => {
                    this.setState({ currentDriveItemId: item.id !== undefined ? item : undefined });
                    this.removeBreadcrumbItems(index + 1);
                  }}>{item.title}</a></li>
                )
              })
            }
          </ul>
          {currentDriveItemId === undefined ?
            <FileList
              pageSize={10}
              disableOpenOnClick
              itemClick={this.itemClick.bind(this)}
              fileListQuery={`/sites/${context.pageContext.site.id}/drive/root/children`}
            >
              <FileTemplate template='file'></FileTemplate>
              <Loading template='loading'></Loading>
              <NoData template='no-data'></NoData>
            </FileList>
            :
            <FileList
              pageSize={10}
              disableOpenOnClick
              fileListQuery={`/me/drives/${currentDriveItemId.parent?.driveId}/items/${currentDriveItemId.id}/children`}
              driveId={currentDriveItemId.parent?.driveId || ""}
              itemId={currentDriveItemId.id}
              itemClick={this.itemClick.bind(this)}
            >
              <FileTemplate template='file'></FileTemplate>
              <Loading template='loading'></Loading>
              <NoData template='no-data'></NoData>
            </FileList>
          }
        </div>
      </section>
    );
  }

  public removeBreadcrumbItems(index: number) {
    const { breadCrumbItems } = this.state;

    this.setState({ breadCrumbItems: breadCrumbItems.slice(0, index) });
  }

  public itemClick(e: CustomEvent<DriveItem>) {
    if (e.detail && e.detail.folder) {
      let id = e.detail.id || "";
      let name = e.detail.name || "";
      let parentReference = e.detail.parentReference || undefined;

      // render new file list
      const { breadCrumbItems } = this.state;

      breadCrumbItems?.push({ title: name, id, parent: parentReference });

      this.setState({ breadCrumbItems, currentDriveItemId: { title: name, id, parent: parentReference } });
    }
    else {
      if (e.detail.webUrl)
        window.open(e.detail.webUrl, "_blank", "noreferrer");
    }
  }

}