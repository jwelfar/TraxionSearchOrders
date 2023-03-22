import * as React from "react";
import { IPrintOrdersProps } from "./IPrintOrdersProps";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPFI, spfi, SPFx } from "@pnp/sp/presets/all";
import {
  DefaultButton,
  TextField,
  ITextFieldStyles,
  Spinner,
  SpinnerSize,
} from "office-ui-fabric-react";
import DataTable from "react-data-table-component";
import * as print from "print-js";

let _sp: SPFI = null;

export interface IDetailsTableItem {
  Title: string;
  NO_ORDEN_REPOSICION_UNOPS: string;
  REGISTRO_SANITARIO: string;
  CANTIDAD_RECIBIDA: string;
}

export interface ITableState {
  columns: any[];
  DatosAI: IDetailsTableItem[];
  NumOrderSearch: string;
  pending: boolean;
  loading: boolean;
}

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
  ITableState
> {
  constructor(props: IPrintOrdersProps) {
    super(props);

    const columnas = [
      {
        id: "column1",
        grow: 2,
        center: true,
        name: "No Orden",
        minWidth: "150px",
        maxWidth: "300px",
        selector: (row: any) => {
          return <span>{row.NO_ORDEN_REPOSICION_UNOPS}</span>;
        },
      },
      {
        id: "column2",
        center: true,
        name: "Cantidad",
        wrap: true,
        minWidth: "150px",
        maxWidth: "300px",
        selector: (row: any) => {
          return <span>{row.CANTIDAD_RECIBIDA}</span>;
        },
      },
      {
        id: "column3",
        center: true,
        name: "Archivo",
        minWidth: "150px",
        maxWidth: "300px",
        selector: (row: any) => {
          const handlePrint = (e: any): any => {
            e.preventDefault();
            print({
              printable: row.LinkTitle,
              type: "pdf",
              showModal: true,
              modalMessage: "Cargando Documento...",
              onError: (error) => {
                alert(`Error found => ${error.message}`);
              },
            });
          };

          return (
            <>
              <DefaultButton text="Imprimir" onClick={(e) => handlePrint(e)} />
            </>
          );
        },
      },
    ];

    this.state = {
      columns: columnas,
      DatosAI: [],
      NumOrderSearch: "",
      pending: true,
      loading: false,
    };
  }

  private async getAIData(): Promise<void> {
    let items: any = [];
    let response: any = [];
    let query = "";
    this.setState({
      loading: true,
    });

    if (this.state.NumOrderSearch.length >= 5) {
      if (query.length === 0) {
        query =
          "substringof('" +
          this.state.NumOrderSearch +
          "', NO_ORDEN_REPOSICION_UNOPS)";
      } else {
        query +=
          " and substringof('" +
          this.state.NumOrderSearch +
          "', NO_ORDEN_REPOSICION_UNOPS)";
      }
    }

    if (this.props.DatosAI) {
      try {
        let next = true;
        if (query.length > 0) {
          items = await getSP(this.props.context)
            .web.lists.getById(this.props.DatosAI.id)
            .items.select(
              "NO_ORDEN_REPOSICION_UNOPS",
              "CANTIDAD_RECIBIDA",
              "Title",
              "LinkTitle"
            )
            .top(50)
            .filter(query)
            .getPaged();
        } else {
          items = await getSP(this.props.context)
            .web.lists.getById(this.props.DatosAI.id)
            .items.select(
              "NO_ORDEN_REPOSICION_UNOPS",
              "CANTIDAD_RECIBIDA",
              "Title",
              "LinkTitle"
            )
            .top(50)
            .getPaged();
        }

        const data = items.results;
        response = response.concat(data);

        while (next) {
          if (items.hasNext) {
            items = await items.getNext();
            response = response.concat(items.results);
          } else {
            next = false;
          }
        }

        this.setState({
          DatosAI: response,
          loading: false,
        });

        return response;
      } catch (err) {
        this.setState({
          loading: false,
        });
        console.log("Error", err);
        err.res.json().then(() => {
          console.log("Failed to get list items!", err);
        });
      }
    }
  }

  handleFilter = async (): Promise<void> => {
    this.setState({
      DatosAI: [],
    });

    setTimeout(async () => {
      await this.getAIData();
    }, 3000);
  };

  async componentDidMount(): Promise<void> {
    await this.getAIData();
    this.setState({
      pending: false,
    });
  }

  public render(): React.ReactElement<IPrintOrdersProps> {
    const textFieldStyles: Partial<ITextFieldStyles> = {
      fieldGroup: { width: 300 },
    };
    return (
      <section>
        <TextField
          label="Buscar por Número de Orden"
          type="search"
          value={this.state.NumOrderSearch}
          onChange={(e) => {
            this.setState(
              {
                NumOrderSearch: (e.target as HTMLInputElement).value,
              },
              async () => {
                await this.handleFilter();
              }
            );
          }}
          styles={textFieldStyles}
        />
        <br />

        {this.state.loading ? (
          <Spinner label="Loading items..." size={SpinnerSize.large} />
        ) : (
          <DataTable
            columns={this.state.columns}
            data={this.state.DatosAI}
            pagination
            progressPending={this.state.pending}
          />
        )}
      </section>
    );
  }
}
