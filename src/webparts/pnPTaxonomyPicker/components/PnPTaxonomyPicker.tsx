import * as React from 'react';
import styles from './PnPTaxonomyPicker.module.scss';
import { IPnPTaxonomyPickerProps } from './IPnPTaxonomyPickerProps';
import { IPnPTaxonomyPickerState } from './IPnPTaxonomyPickerState';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { Checkbox } from 'office-ui-fabric-react/lib/Checkbox';
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';
import { DatePicker } from 'office-ui-fabric-react/lib/DatePicker';
import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { escape } from '@microsoft/sp-lodash-subset';

// @pnp/sp imports    
import { sp } from '@pnp/sp';
import { getGUID } from "@pnp/common";
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { TaxonomyPicker, IPickerTerms } from "@pnp/spfx-controls-react/lib/TaxonomyPicker";
  
// Import button component      
import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react/lib/components/Button'; 
import { autobind } from 'office-ui-fabric-react'; 

export default class PnPTaxonomyPicker extends React.Component<IPnPTaxonomyPickerProps, IPnPTaxonomyPickerState> {
  constructor(props: IPnPTaxonomyPickerProps, state: IPnPTaxonomyPickerState) {  
    super(props);  
    
    this._onRenderFooterContent = this._onRenderFooterContent.bind(this);
    this.state = { 
      addUsers: [], 
      selectedTerms: [], 
      title: '',
      description: '',
      experience: [],
      dpselectedItem: undefined,
      dpselectedItems: [],
      disableToggle: false,
      defaultChecked: false,
      isChecked: false,
      termnCond: false,
      onSubmission: false,
      birthday: null,
      hideDialog: true,
      showPanel: false,
      status: "",
    };  
  }
  
  public render(): React.ReactElement<IPnPTaxonomyPickerProps> {
    const { dpselectedItem } = this.state;
    this._onCheckboxChange = this._onCheckboxChange.bind(this);
    return (
    <form>
      <div className={ styles.pnPTaxonomyPicker }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>!!!Welcome to Employee Registration site!!!</span>
              <p className={ styles.subTitle }>Employee Registration Form Details</p>

              <div className="ms-Grid-col ms-u-sm4 block">
              <label className="ms-Label" style={{color: "blackrgb(51, 51, 51)"}}>EmployeeName:</label>
              </div>
              <div className="ms-Grid-col ms-u-sm8 block">
              <TextField
              required={true}
              type='text'
              onChange={this.myTitleChangeHandler}
              />
              </div>
              <br/>

              <div className="ms-Grid-col ms-u-sm4 block">
              <label className="ms-Label" style={{color: "blackrgb(51, 51, 51)"}}>Job Description:</label>
              </div>
              <div className="ms-Grid-col ms-u-sm8 block">
              <TextField
              multiline={true}
              onChange={this.myDescriptionChangeHandler}
              />
              </div>
              <br/>

              <div className="ms-Grid-col ms-u-sm4 block">
              <label style={{color: "blackrgb(51, 51, 51)"}}>Experience:</label>
              </div>
              <div className="ms-Grid-col ms-u-sm8 block">
              <input
              type='number'
              min={0}
              max={30}
              onChange={this.myExperienceChangeHandler}
              />
              </div>
              <br/>

              <div className="ms-Grid-col ms-u-sm8 block">
                <PeoplePicker
                context={this.props.context}
                titleText="Reporting Manager:"
                personSelectionLimit={3}
                groupName={""} // Leave this blank in case you want to filter from all users
                showtooltip={true}
                isRequired={true}
                disabled={false}
                ensureUser={true}
                selectedItems={this._getPeoplePickerItems}
                showHiddenInUI={false}
                principalTypes={[PrincipalType.User]}
                resolveDelay={1000} />
              </div>
                <br/>

              <div className="ms-Grid-col ms-u-sm8 block">
                <TaxonomyPicker allowMultipleSelections={true}  
                termsetNameOrID="BU"  
                panelTitle="Select Term"  
                label="Project Assigned To:"  
                context={this.props.context}  
                onChange={this.onTaxPickerChange}  
                isTermSetSelectable={false} />
              </div>
                <br/>

              <div className="ms-Grid-col ms-u-sm4 block">
                <label className="ms-Label">Department:</label><br />
              </div>
              <div className="ms-Grid-col ms-u-sm8 block">
                <Dropdown
                placeHolder="Select an Option..."
                label=""
                id="component"
                selectedKey={dpselectedItem ? dpselectedItem.key : undefined}
                ariaLabel="Basic dropdown example"
                options={[
                  { key: 'Human Resource', text: 'Human Resource' },
                  { key: 'Finance', text: 'Finance' },
                  { key: 'Employee', text: 'Employee' }
                ]}
                onChanged={this._changeState}
                onFocus={this._log('onFocus called')}
                onBlur={this._log('onBlur called')}
                />
              </div>
                <br/>

              <div className="ms-Grid-col ms-u-sm4 block">
                <label className="ms-Label">External Hiring:</label>
              </div>
              <div className="ms-Grid-col ms-u-sm8 block">
                <Toggle
                  disabled={this.state.disableToggle}
                  checked={this.state.defaultChecked}
                  label=""
                  onAriaLabel="This toggle is checked. Press to uncheck."
                  offAriaLabel="This toggle is unchecked. Press to check."
                  onText="On"
                  offText="Off"
                  onChanged={(checked) => this._changeSharing(checked)}
                  onFocus={() => console.log('onFocus called')}
                  onBlur={() => console.log('onBlur called')}
                />
              </div>
                <br/>

              <div className="ms-Grid-col ms-u-sm4 block">
                <label className="ms-Label">Date of Birth:</label>
              </div>
              <div className="ms-Grid-col ms-u-sm8 block">
                <DatePicker placeholder="Select a date..."
                  minDate={new Date(1980,12,30)}
                  onSelectDate={this._onSelectDate}
                  value={this.state.birthday}
                  formatDate={this._onFormatDate}
                  isRequired={true}
                />
              </div>
                <br/>

                <Checkbox onChange={this._onCheckboxChange} ariaDescribedBy={'descriptionID'} />
                <span className={`${styles.customFont}`}>I have read and agree to the terms & condition</span><br />
                <p className={(this.state.termnCond === false && this.state.onSubmission === true) ? styles.fontRed : styles.hideElement}>Please check the Terms & Condition</p>
                <br/>

                <PrimaryButton className={ styles.myCustomButton} text="Create" onClick={() => { this.validateForm(); }} />
                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                <PrimaryButton className={ styles.myCustomButton} text="Cancel" onClick={() => { this.setState({}); }} />
                
                <div>
                  <Panel
                    isOpen={this.state.showPanel}
                    type={PanelType.smallFixedFar}
                    onDismiss={this._onClosePanel}
                    isFooterAtBottom={false}
                    headerText="Are you sure you want to create Job Profile ?"
                    closeButtonAriaLabel="Close"
                    onRenderFooterContent={this._onRenderFooterContent}
                  ><span>Please check the details filled and click on Confirm button to create profile.</span>
                  </Panel>
                </div>

                <Dialog
                  hidden={this.state.hideDialog}
                  onDismiss={this._closeDialog}
                  dialogContentProps={{
                    type: DialogType.largeHeader,
                    title: 'Request Submitted Successfully',
                    subText: ""
                  }}
                  modalProps={{
                    titleAriaId: 'myLabelId',
                    subtitleAriaId: 'mySubTextId',
                    isBlocking: false,
                    containerClassName: 'ms-dialogMainOverride'
                  }}>
                  <div dangerouslySetInnerHTML={{ __html: this.state.status }} />
                  <DialogFooter>
                    <PrimaryButton onClick={() => this.gotoHomePage()} text="Okay" />
                  </DialogFooter>
                </Dialog>

            </div>
          </div>
        </div>
      </div>
    </form>
    );
  }

  private myTitleChangeHandler = (event) => {
    this.setState({title: event.target.value});
  }
  private myDescriptionChangeHandler = (event) => {
    this.setState({description: event.target.value});
  }
  private myExperienceChangeHandler = (event) => {
    this.setState({experience: event.target.value});
  }
  private _changeState = (item: IDropdownOption): void => {
    console.log('here is the things updating...' + item.key + ' ' + item.text + ' ' + item.selected);
    this.setState({ dpselectedItem: item });
    if (item.text == "Employee") {
      this.setState({ defaultChecked: false });
      this.setState({ disableToggle: true });
    }
    else {
      this.setState({ disableToggle: false });
    }
  }
  private _log(str: string): () => void {
    return (): void => {
      console.log(str);
    };
  }
  private _changeSharing(checked: any): void {
    this.setState({ defaultChecked: checked });
  }
  private _onCheckboxChange(ev: React.FormEvent<HTMLElement>, isChecked: boolean): void {
    console.log(`The option has been changed to ${isChecked}.`);
    this.setState({ termnCond: (isChecked) ? true : false });
  }
  private _onSelectDate = (date: Date | null | undefined): void => {
    this.setState({ birthday: date });
  }
  private _onFormatDate = (date: Date): string => {
    return date.getDate() + '/' + (date.getMonth() + 1) + '/' + date.getFullYear();
  }
  private _onShowPanel = () => {
    this.setState({ showPanel: true });
  }
  private _onClosePanel = () => {
    this.setState({ showPanel: false });
  }
  private _showDialog = (status: string): void => {
    this.setState({ hideDialog: false });
    this.setState({ status: status });
  }
  private _closeDialog = (): void => {
    this.setState({ hideDialog: true });
  }
  private _onRenderFooterContent = (): JSX.Element => {
    return (
      <div>
        <PrimaryButton onClick={this.addSelectedFields} style={{ marginRight: '8px' }}>
          Confirm
        </PrimaryButton>
        <DefaultButton onClick={this._onClosePanel}>Cancel</DefaultButton>
      </div>
    );
  }
  @autobind 
  private _getPeoplePickerItems(items: any[]) {
    console.log('Items:', items);

    let selectedUsers = [];
    for (let item in items) {
      selectedUsers.push(items[item].id);
    }

    this.setState({ addUsers: selectedUsers });
  }
  @autobind   
  private onTaxPickerChange(terms : IPickerTerms) {  
  console.log("Terms", terms);  
  this.setState({ selectedTerms: terms });  
  }
  /**
   * My sample to show on how form can be validated
   */
  private validateForm(): void {
    let allowCreate: boolean = true;
    this.setState({ onSubmission: true });

    if (this.state.title.length === 0) {
      allowCreate = false;
    }
    if (this.state.selectedTerms === undefined) {
      allowCreate = false;
    }

    if (allowCreate) {
      this._onShowPanel();
    }
    else {
      //do nothing
    }
  }
  @autobind   
  private addSelectedFields(): void {
    this._onClosePanel();
    this._showDialog("Submitting Your Request");
    
    /*
    // Update single value managed metadata field, with first selected term 

    sp.web.lists.getByTitle("SPFxUsers").items.add({  
      //Title: getGUID(),
      Title: this.state.title,
      UsersId: { 
        results: this.state.addUsers
      }, 
      FirstTerm: {   
        __metadata: { "type": "SP.Taxonomy.TaxonomyFieldValue" },  
        Label: this.state.selectedTerms[0].name,  
        TermGuid: this.state.selectedTerms[0].key,  
        WssId: -1  
      }  
    }).then(i => {  
        console.log(i);  
    });   

  */
  
  // Update multi value managed metadata field  
  const spfxList = sp.web.lists.getByTitle('SPFxUsers');    
  // If the name of your taxonomy field is SomeMultiValueTaxonomyField, the name of your note field will be SomeMultiValueTaxonomyField_0  
  const multiTermNoteFieldName = 'Terms_0';  
  let termsString: string = '';  
  this.state.selectedTerms.forEach(term => {  
    termsString += `-1;#${term.name}|${term.key};#`;  
  });  
  
  spfxList.getListItemEntityTypeFullName()  
    .then((entityTypeFullName) => {  
      spfxList.fields.getByTitle(multiTermNoteFieldName).get()  
        .then((taxNoteField) => {  
          const multiTermNoteField = taxNoteField.InternalName;  
          const updateObject = {  
            Title: this.state.title,
            Description: this.state.description,
            UsersId: { 
              results: this.state.addUsers
            },
            Department: this.state.dpselectedItem.key,
            External_x0020_Hiring: this.state.defaultChecked,
            Birthday: this.state.birthday,
            Experience: this.state.experience
          };  
          updateObject[multiTermNoteField] = termsString;  
  
          spfxList.items.add(updateObject, entityTypeFullName)  
            .then((updateResult) => {  
                console.dir(updateResult);
                this.setState({ status: "Your request has been submitted sucessfully." });
            })  
            .catch((updateError) => {  
                console.dir(updateError);  
            });  
        });  
    });   
}
private gotoHomePage(): void {
  window.location.replace(this.props.siteUrl);
}
}
