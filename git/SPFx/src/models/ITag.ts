import { find } from "@microsoft/sp-lodash-subset";
import { ITermInfo } from "@pnp/sp/taxonomy";

export class Term {
    key: string;
    text: string;

    public static mapSPToTerm(props: ITermInfo, locale: string): Term {
        return {
            key: props.id,
            text: find(props.labels, label => label.languageTag === locale)?.name || props.labels[0].name
        };
    }
}