"use strict";
import {IInputs, IOutputs} from "./generated/ManifestTypes";
interface IDropdownOption{
	key:string, text:string
}

export class SelectEntity implements ComponentFramework.StandardControl<IInputs, IOutputs> {

	private notifyOutputChanged: () => void;
	private currentValue: string | undefined;
	private entity:string="";
	private formEntity:string="";
	private container:HTMLDivElement;
	private baseUrl:string;
	private context:ComponentFramework.Context<IInputs>;
	private options: IDropdownOption[]=[];
	private firstRun:Boolean =true;
	private isDisabled:boolean;
	private comboBoxControl: HTMLSelectElement;
	private displayValue:string ="";
	
	constructor()
	{

	}

	/**
	 * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
	 * Data-set values are not initialized here, use updateView.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
	 * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
	 * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
	 * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
	 */
	public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container:HTMLDivElement)
	{
		this.notifyOutputChanged = notifyOutputChanged;
		
		this.container=container;
		this.baseUrl = (<any>context).page.getClientUrl();
		this.context=context;
		let comboBoxContainer = document.createElement("div");
        comboBoxContainer.className = "select-wrapper";

        this.comboBoxControl = document.createElement("select");
        this.comboBoxControl.className = "hdnComboBox";
        this.comboBoxControl.addEventListener("change", this.onChange.bind(this));
        this.comboBoxControl.addEventListener("mouseenter", this.onMouseEnter.bind(this));
		this.comboBoxControl.addEventListener("mouseleave", this.onMouseLeave.bind(this));
		this.comboBoxControl.addEventListener("click", this.onClick.bind(this));

        comboBoxContainer.appendChild(this.comboBoxControl);
        container.appendChild(comboBoxContainer);
	
	}


	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
	public updateView(context: ComponentFramework.Context<IInputs>): void
	{
		this.isDisabled=this.context.mode.isControlDisabled;
		if (this.firstRun){
			this.currentValue=context.parameters.logicalName.raw||"";
			this.firstRun=false;
			this.populateDropdown();
		}		
	}

	/** 
	 * It is called by the framework prior to a control receiving new data. 
	 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
	 */
	public getOutputs(): IOutputs
	{
		return {
			logicalName: this.currentValue,
			displayName: this.displayValue
			};	
	}

	/** 
	 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
	 * i.e. cancelling any pending remote calls, removing listeners, etc.
	 */
	public destroy(): void
	{
		// Add code to cleanup control if necessary
	}

	private async populateDropdown() {

		let selectOption = document.createElement("option");
		var a = await this.getEntities();
		var result = JSON.parse(a);
		var options: IDropdownOption[]=[];
		
		for (var i = 0; i < result.value.length; i++) {
			if (result.value[i].DisplayName != null && result.value[i].DisplayName.UserLocalizedLabel != null) {
				var text = result.value[i].DisplayName.UserLocalizedLabel.Label + " (" + result.value[i].LogicalName + ")";
				var option: IDropdownOption = { key: result.value[i].LogicalName, text: text }
				options.push(option);
				
			}

		}
	
		// sort the items into alphabetical order by text.
		options.sort((a, b) => a.text.localeCompare(b.text));
		if (options.length==1){
			this.currentValue=options[0].key;
		}

		// populate the select option box.

		// firstly remove all existing options.
		for(var i = this.comboBoxControl.options.length - 1 ; i >= 0 ; i--)
		{
			this.comboBoxControl.remove(i);
		}
		
		// add a top level empty option in case it's needed
		
		selectOption = document.createElement("option");
		selectOption.innerHTML="";
		selectOption.value="";
		this.comboBoxControl.add(selectOption);
		
		// add all the sorted records to the list.
		for (let i = 0; i < options.length; i++) {
			
			selectOption = document.createElement("option");
			selectOption.innerHTML = options[i].text;
			selectOption.value = options[i].key;

			if (this.currentValue != null &&
				this.currentValue === options[i].key) {
					selectOption.selected = true;
			//	valueWasChanged = false;
			}

			this.comboBoxControl.add(selectOption);			
		}
		this.comboBoxControl.disabled=this.isDisabled;	
	}

	private async getEntities():Promise<string> {
		var req = new XMLHttpRequest();
		var baseUrl=this.baseUrl;
		
		return new Promise(function (resolve, reject) {

			req.open("GET", baseUrl + "/api/data/v9.1/EntityDefinitions?$select=LogicalName,DisplayName&$filter=IsValidForAdvancedFind%20eq%20true%20and%20IsCustomizable/Value%20eq%20true", true);
			req.onreadystatechange = function () {
				
				if (req.readyState !== 4) return;
				if (req.status >= 200 && req.status < 300) {
					
					// If successful
					try {
						
						var result = JSON.parse(req.responseText);
						if (parseInt(result.StatusCode) < 0) {
							reject({
								status: result.StatusCode,
								statusText: result.StatusMessage
							});
						}
						resolve(req.responseText);
					}
					catch (error) {
						throw error;
					}

				} else {
					// If failed
					reject({
						status: req.status,
						statusText: req.statusText
					});
				}

			};
			req.setRequestHeader("OData-MaxVersion", "4.0");
			req.setRequestHeader("OData-Version", "4.0");
			req.setRequestHeader("Accept", "application/json");
			req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
			req.send();
		});
	}

	private onChange(): void {
		this.currentValue = this.comboBoxControl.value;
		var displayName = this.comboBoxControl.options[this.comboBoxControl.selectedIndex].text;
		displayName=displayName.substring(0, displayName.indexOf(" ("));		
		this.displayValue=displayName;
        this.notifyOutputChanged();
    }

    private onMouseEnter(): void {
        this.comboBoxControl.className = "hdnComboBoxEnter";
	}

    private onMouseLeave(): void {
        this.comboBoxControl.className = "hdnComboBox";
	}
	
	private onClick(): void {
		if (this.comboBoxControl.className="hdnComboBoxFocused"){
		this.comboBoxControl.className = "hdnComboBoxClicked";
		}
		else {
			this.comboBoxControl.className="hdnComboBoxFocused";
		}
	}
	
}