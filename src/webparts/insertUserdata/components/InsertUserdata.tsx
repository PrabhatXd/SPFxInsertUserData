import * as React from 'react';
import type { IInsertUserdataProps } from './IInsertUserdataProps';
import { ITextFieldStyles, TextField, } from '@fluentui/react/lib/TextField';
import { PrimaryButton, mergeStyles } from '@fluentui/react';
import { sp } from '@pnp/sp';
import "@pnp/sp/webs"
import "@pnp/sp/lists"
import "@pnp/sp/items"


const childTextBoxClass = mergeStyles({
  display: 'block',
  marginBottom: '10px',

})
const TextFieldStyles: Partial<ITextFieldStyles> = { root: { maxWidth: '300px', }, }


export default class InsertUserdata extends React.Component<IInsertUserdataProps, {}> {
  public render(): React.ReactElement<IInsertUserdataProps> {


    return (
      <div >
        <h2>Create New Entry to List</h2>
        <TextField
          className={childTextBoxClass}
          styles={TextFieldStyles}
          id="fullname"
          label="Name" />

        <TextField
          className={childTextBoxClass}
          styles={TextFieldStyles}
          id="Department"
          label="Department" />

        <TextField
          className={childTextBoxClass}
          styles={TextFieldStyles}
          id="ReportedTo"
          label="ReportedTo" />

        <TextField
          className={childTextBoxClass}
          styles={TextFieldStyles}
          id="Age"
          label="Age" />

        <TextField
          className={childTextBoxClass}
          styles={TextFieldStyles}
          id="EMPID"
          label="EMPID" />

        <PrimaryButton text='Submit' onClick={this.createItem} />
      </div>
    );
  }
  private createItem = async () => {
    const fullname = (document.getElementById('fullname') as HTMLInputElement).value;
    const Department = (document.getElementById('Department') as HTMLInputElement).value;
    const ReportedTo = (document.getElementById('ReportedTo') as HTMLInputElement).value;
    const Age = (document.getElementById('Age') as HTMLInputElement).value;
    const EMPID = (document.getElementById('EMPID') as HTMLInputElement).value;



    const addItem = await sp.web.lists.getByTitle("CodeAndLearn").items.add({
      'Title': fullname,
      'Department': Department,
      'ReportedTo': ReportedTo,
      'Age': Age,
      'EMPID': EMPID
    })
    console.log(addItem);
    alert(`items created sucessfully with id :${addItem.data.ID}`);
  }
}
