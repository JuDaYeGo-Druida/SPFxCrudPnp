import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './ProyectosWebPart.module.scss';
import * as strings from 'ProyectosWebPartStrings';

import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { IList } from '@pnp/sp/lists';

export interface IProyectosWebPartProps {
  description: string;
}

export default class ProyectosWebPart extends BaseClientSideWebPart<IProyectosWebPartProps> {

  public onInit(): Promise<void> {

    return super.onInit().then(_ => {
      sp.setup({
        spfxContext: this.context
      });
    });

  }

  private AddEventListeners() : void{

    document.getElementById('AddItemToSPList').addEventListener('click',()=>this.AddSPListItem());
    document.getElementById('UpdateItemInSPList').addEventListener('click',()=>this.UpdateSPListItem());
    document.getElementById('DeleteItemFromSPList').addEventListener('click',()=>this.DeleteSPListItem());
    document.getElementsByName('proyectoId').forEach((item) => {
      item.addEventListener('click',()=>this.SelectedItem());
    });
   }

  public render(): void {

    this.domElement.innerHTML = "Cargardo...";
    setTimeout(async () => {
      const items = await sp.web.lists.getByTitle("Proyectos").items.get();

      let html: string = `<div>
      <form>
        <input id="Codigo"  placeholder="Código del Proyecto"/>
        <input id="Titulo"  placeholder="Nombre Proyecto"/>
        <input id="Cliente"  placeholder="Cliente Proyecto"/>
        <input id="Valor"  placeholder="Valor Contrato"/>
        <button id="AddItemToSPList"  type="submit" onclick="return false;" >Añadir</button>
        <button id="UpdateItemInSPList" type="submit"  onclick="return false;" >Actualizar</button>
        <button id="DeleteItemFromSPList"  type="submit"  onclick="return false;" >Eliminar</button>
      </form>
    </div>
    <div id="DivGetItems">
      <table>
      <thead>
        <tr>
          <td>Sel</td>
          <td>Código</td>
          <td>Proyecto</td>
          <td>Cliente</td>
          <td>Valor Contrato</td>
        </tr>
      </thead>`;

      if (items.length > 0) {
        html += `<tbody>`;
        items.forEach((item: any) => {
          html += `<tr>
            <td><input type="radio" name="proyectoId" value="${item.Id}"><br></td>
            <td>${item.C_x00f3_digo_x0020_del_x0020_Pro}</td>
            <td>${item.Title}</td>
            <td>${item.Cliente}</td>
            <td>${item.Valor_x0020_Contrato}</td>
          </tr>`;
        });
        html += `</tbody>`;
      }

      html += `</table><br><br></div><br>`;
      html += `<pre>${JSON.stringify(items, null, 2)}</pre>`;

      this.domElement.innerHTML = html;

      this.AddEventListeners();
    }, 2000);

  }

  private SelectedItem(){
    var proyectoId = this.domElement.querySelector('input[name = "proyectoId"]:checked')["value"];
    sp.web.lists.getByTitle("Proyectos").items.getById(proyectoId).get()
    .then((result: any) => {
      document.getElementById('Codigo')["value"] = result.C_x00f3_digo_x0020_del_x0020_Pro,
      document.getElementById('Titulo')["value"] = result.Title,
      document.getElementById('Cliente')["value"] = result.Cliente,
      document.getElementById('Valor')["value"] = result.Valor_x0020_Contrato
    })
    .catch((error: any) => {
      console.log(proyectoId);
      console.log("Error: " + error);
    });

  }

  private AddSPListItem()
  {
    sp.web.lists.getByTitle("Proyectos").items.add({
      C_x00f3_digo_x0020_del_x0020_Pro : document.getElementById('Codigo')["value"],
      Title : document.getElementById('Titulo')["value"],
      Cliente : document.getElementById('Cliente')["value"],
      Valor_x0020_Contrato : document.getElementById('Valor')["value"]
    })
    .then((result: any) => {
      console.log(result);
      this.render();
      alert("El registro para " + document.getElementById('Codigo')["value"] + " fue añadido!!");
    })
    .catch((error: any) => {
      console.log("Error: " + error);
    });

  }

  private UpdateSPListItem()
  {
    var proyectoId = this.domElement.querySelector('input[name = "proyectoId"]:checked')["value"];
    sp.web.lists.getByTitle("Proyectos").items.getById(proyectoId).update({
      C_x00f3_digo_x0020_del_x0020_Pro : document.getElementById('Codigo')["value"],
      Title : document.getElementById('Titulo')["value"],
      Cliente : document.getElementById('Cliente')["value"],
      Valor_x0020_Contrato : document.getElementById('Valor')["value"]
    })
    .then((result: any) => {
      console.log(result);
      this.render();
      alert("El registro del proyecto con Id " + proyectoId + " fue actualizado!!");
    })
    .catch((error: any) => {
      console.log("Error: " + error);
    });
  }

  private DeleteSPListItem()
  {
    var proyectoId = this.domElement.querySelector('input[name = "proyectoId"]:checked')["value"];
    sp.web.lists.getByTitle("Proyectos").items.getById(proyectoId).delete()
    .then((result: any) => {
      console.log(result);
      this.render();
      alert("El registro del proyecto con Id " + proyectoId + " fue eliminado!!");
    })
    .catch((error: any) => {
      console.log("Error: " + error);
    });
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
