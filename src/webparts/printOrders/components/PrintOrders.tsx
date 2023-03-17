import * as React from "react";
import { IPrintOrdersProps } from "./IPrintOrdersProps";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPFI, spfi, SPFx } from "@pnp/sp/presets/all";
import { TextField } from "office-ui-fabric-react";

let _sp: SPFI = null;

export const getSP = (context?: WebPartContext): SPFI => {
  if (_sp === null && context !== null) {
    //You must add the @pnp/logging package to include the PnPLogging behavior it is no longer a peer dependency
    // The LogLevel set's at what level a message will be written to the console
    _sp = spfi().using(SPFx(context));
  }
  return _sp;
};

export default class PrintOrders extends React.Component<
  IPrintOrdersProps,
  {}
> {
  constructor(props: IPrintOrdersProps) {
    super(props);

    this.state = {};
  }

  public render(): React.ReactElement<IPrintOrdersProps> {
    return (
      <section>
        <TextField
          label="Buscar por Número de Orden"
          type="search"
          // value=""
          // onChange={(e) => {
          //   this.setState({
          //     NumOrderSearch: (e.target as HTMLInputElement).value,
          //   });
          // }}
        />
      </section>
    );
  }
}
