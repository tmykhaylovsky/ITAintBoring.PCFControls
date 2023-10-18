import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { Select2 } from "select2";
import * as $ from 'jquery';
import "./scripts/select2.min.js";

declare let Xrm: any;

class DataAction {
	guid: string;
	associate: boolean;
}
class NToNData {
	len: string;
	ida: string;
	na: string;
	re: string;
	rn: string;
	actions: DataAction[];
}


export class NToNMultiSelect implements ComponentFramework.StandardControl<IInputs, IOutputs> {


	private contextObj: ComponentFramework.Context<IInputs>;
	// Div element created as part of this control's main container
	private mainContainer: HTMLSelectElement;
	private errorElement: HTMLDivElement;
	private selectedItems: string[] = [];
	private overlayDiv: HTMLDivElement;
	private container: HTMLDivElement;
	private _isValidState: boolean = true;

	private _relData: NToNData;

	private _formAttributeLogicalName: string;
	private _linkedEntityName: string;
	private _relationshipEntity: string;
	private _relationshipName: string;
	private _idAttribute: string;
	private _nameAttribute: string;
	// private _linkedEntityFetchXmlResource: string;
	private _linkedEntityDynamicFetchXml: string;
	private _triggerNToNRecordsRefresh: boolean = false;

	private _linkedEntityCollectionName: string;
	private _mainEntityCollectionName: string;

	private _entityMetadataSuccessCallback: any;
	private _linkedEntityMetadataSuccessCallback: any;
	private _relationshipSuccessCallback: any;
	private _successCallback: any;

	private _ctrlId: string;

	private _notifyOutputChanged: () => void;

	private _delayBySeconds: number = 1;


	/**
	 * Empty constructor.
	 */
	constructor() {

	}


	public S4() {
		return (((1 + Math.random()) * 0x10000) | 0).toString(16).substring(1);
	}

	public newGuid() {
		const result: string = (this.S4() + this.S4() + "-" + this.S4() + "-4" + this.S4().substr(0, 3) + "-" + this.S4() + "-" + this.S4() + this.S4() + this.S4()).toLowerCase();
		return result;
	}

	// then to call it, plus stitch in '4' in the third group

	/**
	 * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
	 * Data-set values are not initialized here, use updateView.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
	 * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
	 * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
	 * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
	 */
	public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement) {
		console.log('Starting init(context, notifyOutputChanged, state, container)');

		this.container = container;
		this.contextObj = context;
		if (typeof Xrm == 'undefined') {
			this.errorElement = document.createElement("div");
			this.errorElement.innerHTML = "<H2>This control only works on model-driven forms!</H2>";
			container.appendChild(this.errorElement);
			this._isValidState = false;

			console.log(`Xrm object unavailable`);
		}
		else {

			this._ctrlId = this.newGuid();
			this._relData = new NToNData();
			this._relData.actions = [];

			if (context.parameters.value.attributes?.LogicalName != null) {
				this._formAttributeLogicalName = context.parameters.value.attributes.LogicalName;
			}

			if (context.parameters.linkedEntityName.raw != null) {
				this._linkedEntityName = context.parameters.linkedEntityName.raw;
				this._relData.len = this._linkedEntityName;
			}
			if (context.parameters.idAttribute.raw != null) {
				this._idAttribute = context.parameters.idAttribute.raw;
				this._relData.ida = this._idAttribute;
			}
			if (context.parameters.nameAttribute.raw != null) {
				this._nameAttribute = context.parameters.nameAttribute.raw;
				this._relData.na = this._nameAttribute;
			}
			if (context.parameters.relationshipEntity.raw != null) {
				this._relationshipEntity = context.parameters.relationshipEntity.raw;
				this._relData.re = this._relationshipEntity;
			}
			if (context.parameters.relationshipName.raw != null) {
				this._relationshipName = context.parameters.relationshipName.raw;
				this._relData.rn = this._relationshipName;
			}
			// if (context.parameters.linkedEntityFetchXmlResource.raw != null) {
			// 	this._linkedEntityFetchXmlResource = context.parameters.linkedEntityFetchXmlResource.raw;
			// }

			if (context.parameters.linkedEntityDynamicFetchXml.raw != null) {
				this._linkedEntityDynamicFetchXml = context.parameters.linkedEntityDynamicFetchXml.raw;
			}

			if (context.parameters.triggerNToNRecordsRefresh.raw != null) {
				const triggerRefresh = context.parameters.triggerNToNRecordsRefresh.raw;
				console.log(`triggerRefresh: ${triggerRefresh} (0 = true, 1 = false)`);

				if(context.parameters.triggerNToNRecordsRefresh.raw == '0'){
					this._triggerNToNRecordsRefresh = true;
				}
			}

			console.log(`context.parameters:`);
			console.log(context.parameters);

			context.mode.trackContainerResize(true);
			container.classList.add("pcf_container_element");

			this.overlayDiv = document.createElement("div");
			this.overlayDiv.classList.add("pcf_overlay_element");
			container.appendChild(this.overlayDiv);

			this.mainContainer = document.createElement("select");
			this.mainContainer.id = this._ctrlId;
			this.mainContainer.classList.add("js-example-basic-multiple");
			this.mainContainer.classList.add("pcf_main_element");
			this.mainContainer.multiple = true;
			this.mainContainer.name = "states[]";
			container.appendChild(this.mainContainer);


			this._entityMetadataSuccessCallback = this.entityMetadataSuccessCallback.bind(this);
			this._linkedEntityMetadataSuccessCallback = this.linkedEntityMetadataSuccessCallback.bind(this);
			this._relationshipSuccessCallback = this.relationshipSuccessCallback.bind(this);
			this._successCallback = this.successCallback.bind(this);

			this._notifyOutputChanged = notifyOutputChanged;

			(<any>Xrm).Utility.getEntityMetadata((<any>this.contextObj).page.entityTypeName, []).then(this._entityMetadataSuccessCallback, this.errorCallback);
			(<any>Xrm).Utility.getEntityMetadata(this._linkedEntityName, []).then(this._linkedEntityMetadataSuccessCallback, this.errorCallback);
			//(<any>Xrm).WebApi.retrieveMultipleRecords(this._relationshipEntity, "?$filter="+ (<any>this.contextObj).page.entityTypeName+"id eq " + (<any>this.contextObj).page.entityId, 5000).then(this._relationshipSuccessCallback, this.errorCallback);

			this.getNToNRecordsAndPopulateSelectControlOptions();
			this.attachSelectAndUnselectEvents();
		}
	}

	public entityMetadataSuccessCallback(value: any): void | PromiseLike<void> {
		console.log('Starting entityMetadataSuccessCallback(value)');
		console.log(`value.EntitySetName: ${value.EntitySetName}`);

		this._mainEntityCollectionName = value.EntitySetName;
	}

	public linkedEntityMetadataSuccessCallback(value: any): void | PromiseLike<void> {
		console.log('Starting linkedEntityMetadataSuccessCallback(value)');
		console.log(`value.EntitySetName: ${value.EntitySetName}`);

		this._linkedEntityCollectionName = value.EntitySetName;
	}

	public addOptions(value: any) {
		console.log('Starting addOptions(value)');
		console.log(value);

		for (const i in value.entities) {
			const current: any = value.entities[i];

			const checked = this.selectedItems.indexOf(<string>current[this._idAttribute]) > -1;
			const newOption = new Option(current[this._nameAttribute], current[this._idAttribute], checked, checked);
			$('#' + this._ctrlId).append(newOption);
		}
	}

	public successCallback(value: any): void | PromiseLike<void> {
		console.log('Starting successCallback(value)');
		console.log(value);

		this.addOptions(value);
	}


	public relationshipSuccessCallback(value: any): void | PromiseLike<void> {
		console.log('Starting relationshipSuccessCallback(value)');
		console.log(value);

		if (value != null) {
			for (const i in value.entities) {
				this.selectedItems.push(value.entities[i][this._idAttribute]);
			}
		}
		// if(this._linkedEntityFetchXmlResource != null)
		// {
		// 	var _self = this;
		// 	var xhttp = new XMLHttpRequest();
		// 	xhttp.onreadystatechange = function() {
		// 		if (this.readyState == 4 && this.status == 200) {
		// 		   _self.contextObj.webAPI.retrieveMultipleRecords(_self._linkedEntityName, "?fetchXml=" + encodeURIComponent(this.responseText), 5000).then(_self._successCallback, _self.errorCallback);
		// 		}
		// 	};
		// 	xhttp.open("GET", this._linkedEntityFetchXmlResource, true);
		// 	xhttp.send();

		// }
		// else
		// {
		//     this.contextObj.webAPI.retrieveMultipleRecords(this._linkedEntityName, "?$orderby=" + this._nameAttribute + " asc", 5000).then(this._successCallback, this.errorCallback);
		// }

		const parsedFetchXml = this.parseDynamicFetchXml(this._linkedEntityDynamicFetchXml);
		if (this._linkedEntityDynamicFetchXml != null && parsedFetchXml != null && parsedFetchXml != "") {
			const _self = this;
			_self.contextObj.webAPI.retrieveMultipleRecords(_self._linkedEntityName, "?fetchXml=" + encodeURIComponent(parsedFetchXml), 5000).then(_self._successCallback, _self.errorCallback);
		}
		else {
			this.contextObj.webAPI.retrieveMultipleRecords(this._linkedEntityName, "?$orderby=" + this._nameAttribute + " asc", 5000).then(this._successCallback, this.errorCallback);
		}
	}

	public parseDynamicFetchXml(dynamicFetchXml: string): string {
		console.log('Starting parseDynamicFetchXml(dynamicFetchXml)');
		//console.log(dynamicFetchXml);

		let parsedFetchXml: string = dynamicFetchXml;

		const regex = /{{(.*?)}}/;
		const match = dynamicFetchXml.match(regex);
		console.log(`Checking if FetchXML contains a placeholder to be replaced with an evaluated form lookup attribute value based on regex pattern ${regex}. Match: ${match} (null = no match)`);
		if (match && match.length == 2) {
			const replaceString = match[0];
			const lookupAttributeLogicalName = match[1];

			console.log(`Will replace placeholder ${replaceString} with the id value from a form lookup attribute with logical name of ${lookupAttributeLogicalName}`);

			const lookupAttribute = (<any>Xrm).Page.getAttribute(lookupAttributeLogicalName);

			if (lookupAttribute) {
				const lookupAttributeObject = lookupAttribute.getValue();
				if (lookupAttributeObject) {
					const lookupAttributeId = lookupAttributeObject[0].id.replace('{', '').replace('}', '');

					console.log(`lookupAttributeId: ${lookupAttributeId}`);

					parsedFetchXml = dynamicFetchXml.replace(replaceString, lookupAttributeId);

					console.log(`Parsed FetchXML placeholder (${replaceString}) and replaced it with value (${lookupAttributeId})`);
				} else {
					console.log(`lookupAttributeObject is null (.getValue() returns null) even though lookupAttribute is not null`);
				}
			} else {
				console.log(`lookupAttribute with logical name (${lookupAttributeLogicalName}) doesn't exist on the form. Expected dynamic FetchXML will be invalid`);
			}
		}

		//console.log(`parsedFetchXml: ${parsedFetchXml}`);

		return parsedFetchXml;
	}

	public errorCallback(value: any) {
		console.log('Starting errorCallback(value)');
		console.log(value);

		alert(value);
	}

	public setReadonly(): void {
		console.log(`Starting setReadonly, this.contextObj.mode.isControlDisabled: ${this.contextObj.mode.isControlDisabled} (false = display:none, true = display:block) CSS applied to div.pcf_overlay_element's display property`);

		(<HTMLElement>this.container.firstElementChild).style.display = this.contextObj.mode.isControlDisabled == false ? "none" : "block";
	}

	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
	public updateView(context: ComponentFramework.Context<IInputs>): void {
		console.log('Starting updateView(context) context:');
		console.log(context);
		console.log(`this._isValidState: ${this._isValidState} (true = will attempt to check if triggered to get latest N:N records/rebuild select control)`);

		if (this._isValidState == false) return;
		// Add code to update control view
		this.contextObj = context;
		this.setReadonly();

		const formAttribute = (<any>Xrm).Page.getAttribute(this._formAttributeLogicalName);
		if(formAttribute){
			const formAttributeValue = formAttribute.getValue();

			console.log(`this._triggerNToNRecordsRefresh: ${this._triggerNToNRecordsRefresh} (Trigger N To N Records Refresh PCF control setting)`);
			console.log(`formAttributeValue: ${formAttributeValue} (on init, should be null/empty. on updateView should equal 'trigger' if another script temporarily set value)`);

			if(this._triggerNToNRecordsRefresh && formAttributeValue == 'trigger'){
				console.log(`Remove select and unselect event listeners, otherwise the N:N record create/delete will try to run twice`);
				$('#' + this._ctrlId).select2().off('select2:select');
				$('#' + this._ctrlId).select2().off('select2:unselect');

				console.log(`Clear select control options`);
				$('#' + this._ctrlId).empty();

				console.log(`Get latest N:N records and rebuild select control options. Attach select and unselect events to the select control`);
				this.getNToNRecordsAndPopulateSelectControlOptions();
				this.attachSelectAndUnselectEvents();

				console.log(`Set the triggering form attribute back to null in case it wasn't cleared.`);
				setTimeout(function(){
					formAttribute.setValue(null);
				}, 1000);
			}
		} else {
			console.log(`formAttribute with logical name (${this._formAttributeLogicalName}) doesn't exist on the form`);
		}
		
	}

	public getNToNRecordsAndPopulateSelectControlOptions() : void {
		console.log(`Starting getNToNRecordsAndPopulateSelectControlOptions`);

		if ((<any>this.contextObj).page.entityId != null
			&& (<any>this.contextObj).page.entityId != "00000000-0000-0000-0000-000000000000") {
			this.contextObj.webAPI.retrieveMultipleRecords(this._relationshipEntity, "?$filter=" + (<any>this.contextObj).page.entityTypeName + "id eq " + (<any>this.contextObj).page.entityId, 5000).then(this._relationshipSuccessCallback, this.errorCallback);
		}
		else {
			this.relationshipSuccessCallback(null);
		}
	}

	public attachSelectAndUnselectEvents(): void {
		console.log(`Starting attachSelectAndUnselectEvents()`);

		const thisVar: any = this;
		$(function () {
			thisVar.setReadonly();
			$('#' + thisVar._ctrlId).select2().on('select2:select', function (e) {
				const data = e.params.data;
				thisVar.selectAction("select", data);
			}).on('select2:unselect', function (e) {
				const data = e.params.data;
				thisVar.selectAction("unselect", data);
			});
		});
	}

	/** 
	 * It is called by the framework prior to a control receiving new data. 
	 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
	 */
	public getOutputs(): IOutputs {
		console.log('Starting getOutputs');

		if (this._isValidState == false) {
			return {
				value: ""
			};
		}
		else {
			return {
				value: "NTONDATA:" + JSON.stringify(this._relData)
			};
		}
	}

	/** 
	 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
	 * i.e. cancelling any pending remote calls, removing listeners, etc.
	 */
	public destroy(): void {
		// Add code to cleanup control if necessary
	}

	public selectAction(action: string, data: any) {
		console.log(`Starting selectAction(action: ${action}, data)`);
		console.log(data);

		/*
		function (e: any, data: any) {
					ProcessClick(
					alert("Checked: " + data.node.id);
					alert("Parent: " + data.node.parent); 
					//alert(JSON.stringify(data));
				}
		*/

		if ((<any>this.contextObj).page.entityId == null
			|| (<any>this.contextObj).page.entityId == "00000000-0000-0000-0000-000000000000") {
			if (action == "select") {
				debugger;
				const act = new DataAction();
				act.associate = true;
				act.guid = data.id;
				this._relData.actions.push(act);
			}
			else {
				for (let i = 0; i < this._relData.actions.length; i++) {
					const act = this._relData.actions[i];
					if (act.guid == data.id) {
						this._relData.actions.splice(i, 1);
						break;
					}
				}

			}
			this._notifyOutputChanged();

		}
		else {

			const url: string = (<any>Xrm).Utility.getGlobalContext().getClientUrl();
			const recordUrl: string = url + "/api/data/v9.1/" + this._mainEntityCollectionName + "(" + (<any>this.contextObj).page.entityId + ")";
			const _self = this;
			const thisVar: any = this;

			if (action == "select") {
				Xrm.Utility.showProgressIndicator(`Associating OPS Tag..`);

				//See himbap samples here: http://himbap.com/blog/?p=2063
				const associate = {
					"@odata.id": recordUrl
				};

				const req = new XMLHttpRequest();
				req.open("POST", url + "/api/data/v9.1/" + this._linkedEntityCollectionName + "(" + data.id + ")/" + this._relationshipName + "/$ref", true);
				req.setRequestHeader("Accept", "application/json");
				req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
				req.setRequestHeader("OData-MaxVersion", "4.0");
				req.setRequestHeader("OData-Version", "4.0");
				req.onreadystatechange = function () {
					if (this.readyState == 4 /* complete */) {
						req.onreadystatechange = null;
						if (this.status == 204) {
							//_self.alertDialog('Success',`Successfully associated tag`);
						} else {
							//remove unselect event listener
							$('#' + thisVar._ctrlId).select2().off('select2:unselect');

							//remove item that was added last
							const lastAddedItemSelector = "select#" + thisVar._ctrlId + ` ~ span li[title='${data.text}'] span.select2-selection__choice__remove`;
							console.log(`lastAddedItemSelector: ${lastAddedItemSelector}`);
							jQuery(lastAddedItemSelector).trigger('click');

							//bring back unselect event listener
							$('#' + thisVar._ctrlId).select2().on('select2:unselect', function (e) {
								const data = e.params.data;
								thisVar.selectAction("unselect", data.id);
							});

							const error = JSON.parse(this.response).error;
							_self.alertDialog('Error', error.message);
							_self.showFormNotification(error.message);
						}
					}

					setTimeout(function(){
						Xrm.Utility.closeProgressIndicator();
					}, _self._delayBySeconds * 1000);
				};
				req.send(JSON.stringify(associate));

			}
			else if (action == "unselect") {
				Xrm.Utility.showProgressIndicator(`Disassociating OPS Tag..`);

				//Fixing behavior when select limit reached, the select object changes from a data object containing the id as a property to the data object being the id itself
				if(data.id === undefined){
					data = { id: data };
					console.log('data where data.id was undefined');
					console.log(data);
				}

				const req = new XMLHttpRequest();
				req.open("DELETE", url + "/api/data/v9.1/" + this._linkedEntityCollectionName + "(" + data.id + ")/" + this._relationshipName + "/$ref" + "?$id=" + recordUrl, true);
				req.setRequestHeader("Accept", "application/json");
				req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
				req.setRequestHeader("OData-MaxVersion", "4.0");
				req.setRequestHeader("OData-Version", "4.0");
				req.onreadystatechange = function () {
					if (this.readyState == 4 /* complete */) {
						req.onreadystatechange = null;
						if (this.status == 204) {
							//_self.alertDialog('Success',`Successfully disassociated tag`);
						} else {
							const error = JSON.parse(this.response).error;

							_self.alertDialog('Error', error.message);
							_self.showFormNotification(error.message);
						}
					}

					setTimeout(function(){
						Xrm.Utility.closeProgressIndicator();
					}, _self._delayBySeconds * 1000);
				};
				req.send();
			}
		}

	}

	public showFormNotification(message: string, level: string = 'WARNING', notificationId: any = null) : void{
		console.log('Starting showFormNotification');
		(<any>Xrm).Page.ui.setFormNotification(message, level, notificationId);
	}

	public alertDialog(alertTitle: any, alertText: any) {
		console.log(`Starting alertDialog(alertTitle: '${alertTitle}, alertText: ${alertText})`);

		const alertStrings = { 
			title: alertTitle, 
			text: alertText 
		};
		const alertOptions = { 
			height: 120, 
			width: 260 
		};
		
		(<any>Xrm).Navigation.openAlertDialog(alertStrings, alertOptions).then(
			function success(result: any) {
				console.log("Alert dialog closed");
			},
			function (error: { message: any; }) {
				console.log(error.message);
			}
		);
	}
}

