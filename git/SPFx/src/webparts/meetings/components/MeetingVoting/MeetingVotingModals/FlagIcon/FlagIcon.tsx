import * as React from 'react';
import { IFlagIconprops, IFlagIconState } from './IFlagIcon';
import { Image } from '@fluentui/react-components';
import styles from './FlagIcon.module.scss';

export default class FlagIcon extends React.Component<IFlagIconprops, IFlagIconState> {

    constructor(props: IFlagIconprops) {
        super(props);
    }

    private getFlagImage(communityName: string): string {
        const flagImages: { [key: string]: string } = {
            'Andalucía': require('../assets/Andalucia.png'),
            'Aragón': require('../assets/Aragon.png'),
            'Canarias': require('../assets/Canarias.png'),
            'Cantabria': require('../assets/Cantabria.png'),
            'Castilla y León': require('../assets/CastillaYLeon.png'),
            'Castilla-La Mancha': require('../assets/CastillaLaMancha.png'),
            'Cataluña': require('../assets/Cataluna.png'),
            'Ceuta': require('../assets/Ceuta.png'),
            'Comunidad de Madrid': require('../assets/ComunidadDeMadrid.png'),
            'Comunidad Valenciana': require('../assets/ComunidadValenciana.png'),
            'Extremadura': require('../assets/Extremadura.png'),
            'Galicia': require('../assets/Galicia.png'),
            'Islas Baleares': require('../assets/Baleares.png'),
            'La Rioja': require('../assets/LaRioja.png'),
            'Melilla': require('../assets/Melilla.png'),
            'País Vasco': require('../assets/PaisVasco.png'),
            'Principado de Asturias': require('../assets/Asturias.png'),
            'Región de Murcia': require('../assets/RegionDeMurcia.png'),
            'Comunidad Foral de Navarra': require('../assets/Navarra.png'),
            'España': require('../assets/Espana.png'),
        };

        return (
            flagImages[communityName] || ""
        );
    }

    public render(): React.ReactElement<IFlagIconprops> {
        const { communityName, withShadow } = this.props;

        if (communityName === "Convocante") {
            return (
                <></>
            );
        }

        const flagSrc: string = this.getFlagImage(communityName);
        if (flagSrc === "") {
            return (
                <div> {"Comunidad autónoma no encontrada"} </div>
            );
        }

        return (
            <Image
                className={styles.flagIcon}
                src={flagSrc}
                fit="default"
                shadow={withShadow}
            />
        );
    }

}