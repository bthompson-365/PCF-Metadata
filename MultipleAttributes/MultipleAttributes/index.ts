import {IInputs, IOutputs} from "./generated/ManifestTypes";
import { threadId } from "worker_threads";
interface IMultipleOption{
	key:string, text:string, checked:boolean
}
export class MultipleAttributes implements ComponentFramework.StandardControl<IInputs, IOutputs> {
	/**
	 * Empty constructor.
	 */
	constructor()
	{

	}
	private notifyOutputChanged: () => void;
	private currentValue: string | undefined;
	private dropdownDiv:HTMLDivElement;
	private selectContainer:HTMLDivElement;
	entity:string="";
	container:HTMLDivElement;
	baseUrl:string;
	context:ComponentFramework.Context<IInputs>;
	options: IMultipleOption[]=[];
	firstRun:Boolean =true;
	isDisabled:boolean;
	dropdown:boolean=false;
	currentValues:IMultipleOption[]=[];
	private resultDiv: HTMLDivElement;
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
		var comboBoxContainer = document.createElement("div");
        comboBoxContainer.className = "select-wrapper";
        this.resultDiv = document.createElement("div");
		this.resultDiv.className = "hdnDiv";
		this.resultDiv.id="resultDiv";
        this.resultDiv.addEventListener("click", this.onClick.bind(this));
        this.resultDiv.addEventListener("mouseenter", this.onMouseEnter.bind(this));
        this.container.addEventListener("mouseleave", this.onMouseLeave.bind(this));
		this.dropdownDiv=document.createElement("div");
		this.dropdownDiv.className="hdnDropdown";
		this.selectContainer=document.createElement("div");
		this.selectContainer.className="scrollcontent";
		this.dropdownDiv.appendChild(this.selectContainer);
		comboBoxContainer.appendChild(this.resultDiv);
		comboBoxContainer.appendChild(this.dropdownDiv);
        container.appendChild(comboBoxContainer); 
	}


	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
	public updateView(context: ComponentFramework.Context<IInputs>): void
	{
		this.isDisabled=this.context.mode.isControlDisabled;
		var	entity=context.parameters.EntityName.raw||"";	
		if (entity!==this.entity && entity!==""){
			if (this.firstRun){
				this.currentValue=context.parameters.Attribute.raw||"";
				this.firstRun=false;
			}
			else {
				this.currentValue="";
				this.currentValues=[];
			}
			// time to retrive the actual results from the system.
			this.entity=entity ;		
			this.populateComboBox(entity)	
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
      
        this.notifyOutputChanged();
    }

	private onClick(): void{
		if (!this.context.mode.isControlDisabled){
			this.resultDiv.className="hdnDivClick";
			var width=+this.resultDiv.clientWidth;
			this.dropdownDiv.style.width=(width-2)+"px";
			if (this.selectContainer.hasChildNodes()){
				this.dropdown=!this.dropdown;
				if (this.dropdown){
					this.dropdownDiv.className="hdnDropdownShow";
				}
				else {
					this.resultDiv.className="hdnDivFocused";
					this.dropdownDiv.className="hdnDropdown";
				}
			}
		}
	}

    private onMouseEnter(): void {
		if (this.resultDiv.className==="hdnDiv"){
    	    this.resultDiv.className = "hdnDivFocused";
		}
}

    private onMouseLeave(): void {
		this.resultDiv.className = "hdnDiv";
		this.dropdown=false;
		this.dropdownDiv.className="hdnDropdown";
    }

	private onSelectEnter(div:HTMLDivElement){
		if (!this.context.mode.isControlDisabled){
			div.className="hdnOptionSelected";
		}
	}

	private onSelectLeave(div:HTMLDivElement){
		if (!this.context.mode.isControlDisabled){
			var result:boolean=false;
			var elements=div.getElementsByTagName("input");
			if (elements.length>0){
				var c=<HTMLInputElement>elements[0];
				if (c!==undefined && c.checked){
					result=true;
				}

			}
			if (!result){
				div.className="hdnOption";
			}
		}
	}

	private findOption(key:string):IMultipleOption{
		var option: IMultipleOption = { key: "", text: "", checked:false };
		for (var i=0; i<this.options.length; i++){
			if (this.options[i].key==key){
				option=this.options[i];
			}
		}	
		return option;
	}

	private onSelectClick(div:HTMLDivElement){
		if (!this.context.mode.isControlDisabled){
			var elements=div.getElementsByTagName("input");
			if (elements.length>0){
				var c=<HTMLInputElement>elements[0];
				if (c!==undefined){ 
					c.checked=!c.checked;

				}
				var item=this.findOption(c.value);
				if (item.key!==""){
				if (c.checked){
					item.checked=true;
					this.currentValues.push(item);
					div.className="hdnOptionSelected";
				} else {
					this.currentValues.splice( this.currentValues.indexOf(item),1);
					// change font to emphasis that it has been clicked
					div.className="hdnOption";
				}
			}
			this.currentValues.sort((a, b) => a.text.localeCompare(b.text));
			var string="";
			this.currentValue="";
			for (var i=0; i<this.currentValues.length; i++){
				string=string+this.currentValues[i].text+", ";
				this.currentValue=this.currentValue+this.currentValues[i].key+", ";
			}
			this.resultDiv.innerHTML=string.slice(0,-2);
			this.currentValue=this.currentValue.slice(0,-2);
			this.notifyOutputChanged();
			}
		}
	}


	private async populateComboBox(entity:string) {

		let optionDiv = document.createElement("div");
		if (entity!=""){
			var a = await this.getAttributes(entity);
			var result = JSON.parse(a);
			var options: IMultipleOption[]=[];
			var selectedItems:string[]=[];
			if (this.currentValue!=="" && this.currentValue!==undefined){
				selectedItems=this.currentValue.split(", ");
			}
			// format all the options into a usable record
			for (var i = 0; i < result.value.length; i++) {
				
				if (result.value[i].DisplayName != null && result.value[i].DisplayName.UserLocalizedLabel != null) {
					var key=result.value[i].LogicalName;
					var checked:boolean=false;
					var text = result.value[i].DisplayName.UserLocalizedLabel.Label + " (" + result.value[i].LogicalName + ")";
					if (selectedItems.includes(key)){
						checked=true;
						var option: IMultipleOption = { key: key, text: text, checked:checked }
						this.currentValues.push(option);
					}
					
					var option: IMultipleOption = { key: key, text: text, checked:checked }
					options.push(option);
					
				}
			}
			// sort the items into alphabetical order by text.
			options.sort((a, b) => a.text.localeCompare(b.text));
			this.options=options;
			// populate the select option box.
			for (var i= this.selectContainer.children.length; i>0; i--){
				this.selectContainer.removeChild(this.selectContainer.children[i-1]);
			}
			

			// add a top level empty option in case it's needed
		//	optionDiv = document.createElement("div");
			
		//	optionDiv.innerHTML="&nbsp;";
		//	selectOption.value="";
		//	this.selectContainer.appendChild(optionDiv);
			

			// add all the sorted records to the list.
			for (let i = 0; i < options.length; i++) {
				
				optionDiv = document.createElement("div");
				optionDiv.className="hdnOption";
				var select:HTMLInputElement=document.createElement("input");
				select.type="Checkbox";
				select.name=options[i].text;
				select.value=options[i].key;
				select.className="hdnCheckbox";
				if (options[i].checked){
					select.checked=true;
					optionDiv.className = "hdnOptionSelected";
				}
				
				var span:HTMLSpanElement= document.createElement("span");
				// create checkbox
				span.innerHTML = options[i].text;
				//selectOption.value = options[i].key;

				if (this.currentValue != null &&
					this.currentValue === options[i].key) {
						select.checked=true;
				//	valueWasChanged = false;
				}
				optionDiv.appendChild(select);
				optionDiv.appendChild(span);
				optionDiv.addEventListener("mouseenter", this.onSelectEnter.bind(this, optionDiv));
				optionDiv.addEventListener("mouseleave", this.onSelectLeave.bind(this,optionDiv));
				optionDiv.addEventListener("click",this.onSelectClick.bind(this,optionDiv));
				this.selectContainer.appendChild(optionDiv);
				
			}
		}
		
		//this.dropDownControl.disabled=this.isDisabled;
		var string="";
		this.currentValue="";
		for (i=0; i<this.currentValues.length; i++){
			string=string+this.currentValues[i].text+", ";
			this.currentValue=this.currentValue+this.currentValues[i].key+", ";
		}
		this.resultDiv.innerHTML=string.slice(0,-2);
		this.currentValue=this.currentValue.slice(0,-2);
		
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