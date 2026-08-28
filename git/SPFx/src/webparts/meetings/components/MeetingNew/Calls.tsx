import * as React from 'react';
import styles from './Calls.module.scss';
import { ICallsState } from './ICallsState';
import { ICallsprops } from './ICallsprops';
import * as strings from 'MeetingsWebPartStrings';
import {
  Add20Filled
} from '@fluentui/react-icons';
import {
  Button
} from '@fluentui/react-components';
import MeetingForm from '../../../../components/MeetingForm';
import { IEvent } from '../../../../service/BackendServiceModels/EventModels';

export default class Calls extends React.Component<ICallsprops, ICallsState> {
  constructor(props: ICallsprops) {
    super(props);
    this.state = {
      openForm: false
    }
  }

  public openFormFromParent() {
    this.setState({ openForm: !this.state.openForm });
  }

  private onChangedEvent(event: IEvent) {
    //realizar accion necesaria para despues de la carga (por ejemplo esperar a por la creacion con un boton de cargando)
  }

  public render(): React.ReactElement<ICallsprops> {
    return (
      <section className={styles.calls}>
        <Button className={styles.mainButton} onClick={() => {
          this.setState({ openForm: !this.state.openForm });
        }} icon={<Add20Filled />} appearance='primary'>{strings.CreateConvo}</Button>
        {
          this.state.openForm &&
          <MeetingForm isEditor={true} {...this.props} readOnly={false} editMode={false} dataChanged={this.onChangedEvent} openForm={this.openFormFromParent.bind(this)}></MeetingForm>
        }
      </section>
    );
  }

}