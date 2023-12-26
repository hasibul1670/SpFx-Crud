import { sp } from "@pnp/sp";
import "@pnp/sp/items";
import "@pnp/sp/lists";
import "@pnp/sp/webs";
import * as React from "react";
import styles from "./Crud.module.scss";
import type { ICrudProps } from "./ICrudProps";

export default class Crud extends React.Component<ICrudProps, {}> {
  public render(): React.ReactElement<ICrudProps> {
    return (
      <div className={styles.spfxCrudPnp}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <div className={styles.itemField}>
                <div className={styles.fieldLabel}>Item ID:</div>
                <input type="text" id="itemId"></input>
              </div>
              <div className={styles.itemField}>
                <div className={styles.fieldLabel}>Full Name</div>
                <input type="text" id="fullName"></input>
              </div>
              <div className={styles.itemField}>
                <div className={styles.fieldLabel}>Age</div>
                <input type="text" id="age"></input>
              </div>
              <div className={styles.itemField}>
                <div className={styles.fieldLabel}>All Items:</div>
                <div id="allItems"></div>
              </div>
              <div className={styles.buttonSection}>
                <div className={styles.button}>
                  <span className={styles.label} onClick={this.createItem}>
                    Create
                  </span>
                </div>
                <div className={styles.button}>
                  <span className={styles.label} onClick={this.getItemById}>
                    Read By Id
                  </span>
                </div>
                <div className={styles.button}>
                  <span className={styles.label} onClick={this.getAllItems}>
                    Read All
                  </span>
                </div>
                <div className={styles.button}>
                  <span className={styles.label} onClick={this.updateItem}>
                    Update
                  </span>
                </div>
                <div className={styles.button}>
                  <span className={styles.label} onClick={this.deleteItem}>
                    Delete
                  </span>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }

  // Create Item
  private createItem = async () => {
    try {
      const fullNameInput: HTMLInputElement | null = document.getElementById(
        "fullName"
      ) as HTMLInputElement;
      const ageInput: HTMLInputElement | null = document.getElementById(
        "age"
      ) as HTMLInputElement;

      if (fullNameInput && ageInput) {
        const addItem = await sp.web.lists
          .getByTitle("EmployeeDetails")
          .items.add({
            Title: fullNameInput.value,
            Age: ageInput.value,
          });

        console.log(addItem);
        alert(`Item created successfully with ID: ${addItem.data.ID}`);
      } else {
        console.error("Input elements not found.");
      }
    } catch (e) {
      console.error(e);
    }
  };

  // Get Item by ID
  private getItemById = async () => {
    try {
      const idInput: HTMLInputElement | null = document.getElementById(
        "itemId"
      ) as HTMLInputElement;

      if (idInput?.value) {
        const id: number = parseInt(idInput.value);

        if (id > 0) {
          const item: any = await sp.web.lists
            .getByTitle("EmployeeDetails")
            .items.getById(id)
            .get();

          const fullNameInput: HTMLInputElement | null =
            document.getElementById("fullName") as HTMLInputElement;
          const ageInput: HTMLInputElement | null = document.getElementById(
            "age"
          ) as HTMLInputElement;

          if (fullNameInput && ageInput) {
            fullNameInput.value = item.Title;
            ageInput.value = item.Age;
          } else {
            console.error("Input elements not found.");
          }
        } else {
          alert(`Please enter a valid item id.`);
        }
      } else {
        alert(`Please enter a valid item id.`);
      }
    } catch (e) {
      console.error(e);
    }
  };

  // Get all items
  private getAllItems = async () => {
    console.log(sp.web.lists.getByTitle("EmployeeDetails").items.get());
    try {
      const items: any[] = await sp.web.lists
        .getByTitle("EmployeeDetails")
        .items.get();

      console.log(items);

      const allItemsElement: HTMLElement | null =
        document.getElementById("allItems");

      if (allItemsElement) {
        if (items.length > 0) {
          var html = `<table><tr><th>ID</th><th>Full Name</th><th>Age</th></tr>`;
          items.map((item, index) => {
            html += `<tr><td>${item.ID}</td><td>${item.Title}</td><td>${item.Age}</td></li>`;
          });
          html += `</table>`;
          allItemsElement.innerHTML = html;
        } else {
          allItemsElement.innerHTML = "List is empty.";
        }
      } else {
        console.error("Element with ID 'allItems' not found.");
      }
    } catch (e) {
      console.error(e.message);
    }
  };

  // Update Item
  private updateItem = async () => {
    try {
      const idInput: HTMLInputElement | null = document.getElementById(
        "itemId"
      ) as HTMLInputElement;

      if (idInput?.value) {
        const id: number = parseInt(idInput.value);

        if (id > 0) {
          const fullName: string =
            (document.getElementById("fullName") as HTMLInputElement)?.value ||
            "";
          const age: number = parseInt(
            (document.getElementById("age") as HTMLInputElement)?.value || "0"
          );

          const itemUpdate = await sp.web.lists
            .getByTitle("EmployeeDetails")
            .items.getById(id)
            .update({
              Title: fullName,
              Age: age,
            });

          console.log(itemUpdate);
          alert(`Item with ID: ${id} updated successfully!`);
        } else {
          alert(`Please enter a valid item id.`);
        }
      } else {
        alert(`Please enter a valid item id.`);
      }
    } catch (e) {
      console.error(e);
    }
  };

  // Delete Item
  private deleteItem = async () => {
    try {
      const idInput: HTMLInputElement | null = document.getElementById(
        "itemId"
      ) as HTMLInputElement;

      if (idInput?.value) {
        const id: number = parseInt(idInput.value);

        if (id > 0) {
          let deleteItem = await sp.web.lists
            .getByTitle("EmployeeDetails")
            .items.getById(id)
            .delete();
          console.log(deleteItem);
          alert(`Item ID: ${id} deleted successfully!`);
        } else {
          alert(`Please enter a valid item id.`);
        }
      } else {
        alert(`Please enter a valid item id.`);
      }
    } catch (e) {
      console.error(e);
    }
  };
}
