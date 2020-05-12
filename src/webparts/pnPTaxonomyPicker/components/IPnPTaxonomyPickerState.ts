import { IPickerTerms } from "@pnp/spfx-controls-react/lib/TaxonomyPicker";
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/components/Dropdown';  
  
export interface IPnPTaxonomyPickerState {  
    selectedTerms: IPickerTerms;
    addUsers: number[];
    title: string;
    description: string;
    experience: number[];
    dpselectedItem?: { key: string | number | undefined };
    dpselectedItems: IDropdownOption[];
    disableToggle: boolean;
    defaultChecked: boolean;
    isChecked: boolean;
    termnCond:boolean;
    onSubmission:boolean;
    birthday?:any|null;
    hideDialog: boolean;
    showPanel: boolean;
    status: string;
}