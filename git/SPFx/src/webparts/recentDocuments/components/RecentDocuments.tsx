import * as React from 'react';
import * as strings from 'RecentDocumentsWebPartStrings';
import styles from './RecentDocuments.module.scss';
import { IRecentDocumentsprops } from './IRecentDocumentsprops';
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
    <div className={styles.spinnerContainer}>
      <Spinner style={{ flex: 1 }} label={strings.LoadingRecentDocuments + "..."} />
    </div>
  );
}

export const NoData = (props: MgtTemplateprops) => {
  return (
    <span className={styles.noDocumentsAvailable} >{strings.NoDocumentsAvailable}</span>
  );
}
export default class RecentDocuments extends React.Component<IRecentDocumentsprops, {}> {

  constructor(props: IRecentDocumentsprops) {
    super(props);
  }

  public renderMoreLink = (): React.ReactElement => {
    return <a className={styles.moreLink} href={this.props.moreLink} title={strings.ViewMoreRecentDocuments}> {strings.MoreLink} </a>;
  };

  public render(): React.ReactElement<IRecentDocumentsprops> {
    const { title, moreLink } = this.props;

    return (
      <section className={styles.recentDocuments}>
        <div className={styles.titleRow}>
          <div className={styles.wpTitle}>{title}</div>
          {
            moreLink && this.renderMoreLink()
          }
        </div>
        <div className={styles.mainContainer}>
          <FileList
            pageSize={10}
            fileListQuery={`/me/drive/recent`}
          >
            <FileTemplate template='file'></FileTemplate>
            <Loading template='loading'></Loading>
            <NoData template='no-data'></NoData>
          </FileList>
        </div>
      </section>
    );
  }

}