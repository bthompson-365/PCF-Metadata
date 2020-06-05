import {IInputs, IOutputs} from "./generated/ManifestTypes";
interface IDropdownOption{
	key:string, text:string
}

export class SelectAttribute implements ComponentFramework.StandardControl<IInputs, IOutputs> {
	
	/**
	 * Empty constructor.
	 */
	constructor()
	{

		
	}
	private notifyOutputChanged: () => void;
	private currentValue: string | undefined;
	private allowedTypes : string;
	entity:string="";
	container:HTMLDivElement;
	baseUrl:string;
	context:ComponentFramework.Context<IInputs>;
	options: IDropdownOption[]=[];
	firstRun:Boolean =true;
	isDisabled:boolean;
	private comboBoxControl: HTMLSelectElement;
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
		// Add control initialization code
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
		
		if(context.parameters.allowedTypes.raw != null){
			this.allowedTypes = context.parameters.allowedTypes.raw;
		}
	}
	
	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
	public updateView(context: ComponentFramework.Context<IInputs>): void
	{
		this.isDisabled=this.context.mode.isControlDisabled;
		var	entity=context.parameters.Entity.raw||"";
		
		if (entity!==this.entity && entity!==""){
			if (this.firstRun){
				this.currentValue=context.parameters.Attribute.raw||"";
				this.firstRun=false;
			}
			else {
				this.currentValue="";
			}
			// time to retrive the actual results from the system.
			this.entity=entity ;		
			this.populateComboBox(entity);	
		}
	}

	/** 
	 * It is called by the framework prior to a control receiving new data. 
	 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
	 */
	public getOutputs(): IOutputs
	{	
		return {
			Attribute:this.currentValue
		}
	}

	/** 
	 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
	 * i.e. cancelling any pending remote calls, removing listeners, etc.
	 */
	public destroy(): void
	{
		// Add code to cleanup control if necessary
	}

	private onChange(): void {
        this.currentValue = this.comboBoxControl.value;
        this.notifyOutputChanged();
    }

    private onMouseEnter(): void {
        this.comboBoxControl.className = "hdnComboBoxFocused";
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

	private async populateComboBox(entity:string) {
		let selectOption = document.createElement("option");
		if (entity!==""){
			var a = await this.getAttributes(entity);
			var result = JSON.parse(a);
			var options: IDropdownOption[]=[];
			// format all the options into a usable record
			for (var i = 0; i < result.value.length; i++) {
				
				if (result.value[i].DisplayName !== null && 
					result.value[i].DisplayName.UserLocalizedLabel !== null && 
					this.isAllowedType(result.value[i].AttributeType)) {
					
					var text = result.value[i].DisplayName.UserLocalizedLabel.Label + " (" + result.value[i].LogicalName + ")";
					var option: IDropdownOption = { key: result.value[i].LogicalName, text: text }
					options.push(option);
					
				}
			}
			// sort the items into alphabetical order by text.
			options.sort((a, b) => a.text.localeCompare(b.text));
			
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

				if (this.currentValue !== null &&
					this.currentValue === options[i].key) {
						selectOption.selected = true;
				//	valueWasChanged = false;
				}

				this.comboBoxControl.add(selectOption);
				
			}
		}		
		this.comboBoxControl.disabled=this.isDisabled;		
	}

	private isAllowedType(type: string): boolean{
		if(this.allowedTypes === "All"){
			return true;
		}else{
			return this.allowedTypes.indexOf(type) > -1
		}
	}

	private async getAttributes(entity: string):Promise<string> {
		var req = new XMLHttpRequest();
		var baseUrl=this.baseUrl;
		return new Promise(function (resolve, reject) {

			req.open("GET", baseUrl + "/api/data/v9.1/EntityDefinitions(LogicalName='"+entity+"')/Attributes?$select=LogicalName,DisplayName,AttributeType&$filter=AttributeOf%20eq%20null", true);
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
}