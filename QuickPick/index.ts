import {IInputs, IOutputs} from "./generated/ManifestTypes";

export class QuickPick implements ComponentFramework.StandardControl<IInputs, IOutputs> {

    private _RemoveFromQueue: boolean = false;
    private _container: HTMLDivElement;
    private _context: ComponentFramework.Context<IInputs>;
    private _userId: string;
    private _currentEntityName: string;
    private _currentEntityId: string;    
    private _controlViewRendered: Boolean;
    private _resultDivContainer: HTMLDivElement;

    //private notifyOutputChanged: () => void;

	/**
	 * Empty constructor.   
	 */
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
    public init(context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary,
        container: HTMLDivElement) {
        this._context = context;
        this._controlViewRendered = false;
        this._container = document.createElement("div");
        this._container.classList.add("QuickPick_Container");
        container.appendChild(this._container);
        //this._RemoveFromQueue = context.parameters.RemoveFromQueue.raw || false;//struggling to take enum parameter

        //@ts-ignore
        this._currentEntityId = this._context.mode.contextInfo.entityId;
        //@ts-ignore
        this._currentEntityName = this._context.mode.contextInfo.entityTypeName;
        this._userId = this._context.userSettings.userId;
	}

	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
	public updateView(context: ComponentFramework.Context<IInputs>): void
	{
        if (!this._controlViewRendered) {
            this._controlViewRendered = true;
            
            this.renderButtonDiv();
            this.renderResultsDiv();
        }
	}

	/** 
	 * It is called by the framework prior to a control receiving new data. 
	 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
	 */
	public getOutputs(): IOutputs
	{
		return {};
	}

	/** 
	 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
	 * i.e. cancelling any pending remote calls, removing listeners, etc.
	 */
	public destroy(): void
	{
		// Add code to cleanup control if necessary
    }

    private renderButtonDiv() {
        let thisRef = this;

        this._context.webAPI.retrieveMultipleRecords("queueitem", "?$filter=_objectid_value eq " + thisRef._currentEntityId + " and statecode eq 0&$orderby=modifiedon desc&$select=queueitemid,_workerid_value").then(
            function (response: any) {
                response.entities.forEach((qi: any) => {
                    let createQuickPickButton = thisRef.createHTMLButtonElement(
                        "Pick This Record",
                        "pickButton-" + qi.queueitemid,
                        qi.queueitemid,
                        thisRef.createPickButtonOnClickHandler.bind(thisRef)
                    );

                    thisRef._container.appendChild(createQuickPickButton);

                    let createQuickReleaseButton = thisRef.createHTMLButtonElement(
                        "Release This Record",
                        "releaseButton-" + qi.queueitemid,
                        qi.queueitemid,
                        thisRef.createReleaseButtonOnClickHandler.bind(thisRef)
                    );

                    thisRef._container.appendChild(createQuickReleaseButton);

                    thisRef.getCurrentWorker(response);
                    console.log("hello");
                });
            },
            function (errorResponse: any) {
                thisRef.updateResultContainerTextWithErrorResponse(errorResponse);
            }
        );
    }

    private async getCurrentWorker(response: any) {
        let thisRef = this;
        if (response === null) {
            let queueItems = await this._context.webAPI.retrieveMultipleRecords("queueitem", "?$filter=_objectid_value eq " + thisRef._currentEntityId + " and statecode eq 0&$orderby=modifiedon desc&$select=queueitemid,_workerid_value");

            queueItems.entities.forEach(async (qi: any) => {
                if (qi._workerid_value === null) {
                    thisRef.updateResultContainerText("No one is working on this record");
                } else {
                    let name = await thisRef.getUserNameById(qi._workerid_value);
                    thisRef.updateResultContainerText(name + " is working on this record");
                }
            });
        } else {
            response.entities.forEach((qi: any) => {
                if (qi._workerid_value === null) {
                    thisRef.updateResultContainerText("No one is working on this record");
                } else {
                    thisRef.getUserNameById(qi._workerid_value).then(
                        success => {
                            thisRef.updateResultContainerText(name + " is working on this record");                            
                    }, failure => { });
                }
            });
        }
    }

    private async getUserNameById(systemUserId: string): Promise<string> {
        let thisRef = this;
        let name = "";
        try {

            let response = await this._context.webAPI.retrieveRecord("systemuser", systemUserId, "?$select=fullname");
            name = response.fullname;
        } catch (err) {
            console.log(err);
        }
        return name;
    }

    /**
     * Function to get associated team members for routing.
     * @param systemUserId
     */
    private async getTeamMembers(systemUserId: string): Promise<any> {
        let thisRef = this;
        let name = "";
        let teamMembersQuery:string = '<fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="true" >' +
            '<entity name="systemuser" >' + 
        '<attribute name="fullname" />' + 
            '<attribute name="businessunitid" />' + 
            '<attribute name="title" />' + 
            '<attribute name="address1_telephone1" />' +
            '<attribute name="positionid" />' +
            '<attribute name="emailaddress" />' + 
            '<attribute name="systemuserid" />' + 
            '<order attribute="fullname" descending="false" />' + 
            '<filter type="and" >' +
            '<condition attribute="systemuserid" operator="neq" uitype="systemuser" value="' + systemUserId + '" />' +
            '</filter>' + 
            '<link-entity name="teammembership" from="systemuserid" to="systemuserid" visible="false" intersect="true" >' + 
        '<link-entity name="team" from="teamid" to="teamid" alias="af" >' + 
            '<filter>' + 
            '<condition attribute="isdefault" operator="eq" value="0" />' + 
            '</filter>' + 
        '<link-entity name="teammembership" from="teamid" to="teamid" visible="false" intersect="true" >' + 
            '<link-entity name="systemuser" from="systemuserid" to="systemuserid" alias="ag" >' + 
        '<filter type="and" >' + 
            '<condition attribute="systemuserid" operator="eq" uitype="systemuser" value="' + systemUserId + '" />' + 
         '</filter>' + 
        '</link-entity>' + 
        '</link-entity>' + 
                 '</link-entity>' + 
        '</link-entity>' + 
             ' </entity>' + 
            '</fetch>';

        try {
            let results = await this._context.webAPI.retrieveMultipleRecords("systemuser", "?fetchXml=" + teamMembersQuery);
            return results;
        } catch (err) {
            console.log(err);
        }
    }

    private renderResultsDiv() {
        this._resultDivContainer = this.createHTMLDivElement(
            "result_container",
            false,
            undefined
        );
        
        this._container.appendChild(this._resultDivContainer);
        // Init the result container with a notification the control was loaded
        this.updateResultContainerText("Control Loaded Correctly");
    }

    /**
   * Helper method to create HTML Div Element
   * @param elementClassName : Class name of div element
   * @param isHeader : True if 'header' div - adds extra class and post-fix to ID for header elements
   * @param innerText : innerText of Div Element
   */
    private createHTMLDivElement(
        elementClassName: string,
        isHeader: boolean,
        innerText?: string): HTMLDivElement {
        let div: HTMLDivElement = document.createElement("div");
        if (isHeader) {
            div.classList.add("QuickPick_Header");
            elementClassName += "_header";
        }
        if (innerText) {
            div.innerText = innerText.toUpperCase();
        }
        div.classList.add(elementClassName);
        return div;
    }

    /**
   * Helper method to inject HTML into result container div
   * @param message : HTML to inject into result container
   */
    private updateResultContainerText(message: string): void {
        if (this._resultDivContainer) {
            this._resultDivContainer.innerHTML = message;
        }
    }
    /**
     * Helper method to inject error string into result container div after failed Web API call
     * @param errorResponse : error object from rejected promise
     */
    private updateResultContainerTextWithErrorResponse(errorResponse: any): void {
        if (this._resultDivContainer) {
            // Retrieve the error message from the errorResponse and inject into the result div
            let errorHTML: string = "Error with this control:";
            errorHTML += "<br />";
            errorHTML += errorResponse.message;
            this._resultDivContainer.innerHTML = errorHTML;
        }
    }

    /**
   * Event Handler for onClick of button
   * @param event : click event
   */
    private createPickButtonOnClickHandler(event: Event): void {
        var thisRef = this;

        let queueItemId: string = (<HTMLInputElement>event.target!).attributes.getNamedItem("buttonvalue")!.value;

        this.PickRecord(queueItemId);
    }

    /**
   * Event Handler for onClick of button
   * @param event : click event
   */
    private createReleaseButtonOnClickHandler(event: Event): void {
        var thisRef = this;

        let queueItemId: string = (<HTMLInputElement>event.target!).attributes.getNamedItem("buttonvalue")!.value;

        thisRef.ReleaseRecord(queueItemId);
    }

    private PickRecord(queueItemId: string): void {
        let thisRef = this;
        
        var parameters: any = {};
        var entity: any = {};
        entity.id = queueItemId;
        entity.entityType = "queueitem";
        parameters.entity = entity;
        var systemuser: any = {};
        systemuser.systemuserid = thisRef._userId; 
        systemuser["@odata.type"] = "Microsoft.Dynamics.CRM.systemuser";
        parameters.SystemUser = systemuser;
        parameters.RemoveQueueItem = thisRef._RemoveFromQueue;

        var pickFromQueueRequest = {
            entity: parameters.entity,
            SystemUser: parameters.SystemUser,
            RemoveQueueItem: parameters.RemoveQueueItem,

            getMetadata: function () {
                return {
                    boundParameter: "entity",
                    parameterTypes: {
                        "entity": {
                            "typeName": "mscrm.queueitem",
                            "structuralProperty": 5
                        },
                        "SystemUser": {
                            "typeName": "mscrm.systemuser",
                            "structuralProperty": 5
                        },
                        "RemoveQueueItem": {
                            "typeName": "Edm.Boolean",
                            "structuralProperty": 1
                        }
                    },
                    operationType: 0,
                    operationName: "PickFromQueue"
                };
            }
        };

        if (pickFromQueueRequest) {
            // @ts-ignore
            Xrm.WebApi.online.execute(pickFromQueueRequest).then(
                function success(result: any) {
                    if (result.ok) {
                        thisRef.getCurrentWorker(null);
                    }
                },
                function (error: any) {
                    thisRef.updateResultContainerText("an error occured");
                }
            );
        }
    }

    private ReleaseRecord(queueItemId: string): void {
        let thisRef = this;

        var parameters: any = {};
        var entity: any = {};
        entity.id = queueItemId;
        entity.entityType = "queueitem";
        parameters.entity = entity;

        var releaseFromQueueRequest = {
            entity: parameters.entity,

            getMetadata: function () {
                return {
                    boundParameter: "entity",
                    parameterTypes: {
                        "entity": {
                            "typeName": "mscrm.queueitem",
                            "structuralProperty": 5
                        }
                    },
                    operationType: 0,
                    operationName: "ReleaseToQueue"
                };
            }
        };

        if (releaseFromQueueRequest) {
            // @ts-ignore
            Xrm.WebApi.online.execute(releaseFromQueueRequest).then(
                function success(result: any) {
                    if (result.ok) {
                        thisRef.getCurrentWorker(null);
                    }
                },
                function (error: any) {
                    thisRef.updateResultContainerText("Ultra fail n00b");
                }
            );
        }
    }


    /**
 * Helper method to create HTML Button used for CreateRecord Web API Example
 * @param buttonLabel : Label for button
 * @param buttonId : ID for button
 * @param buttonValue : value of button (attribute of button)
 * @param onClickHandler : onClick event handler to invoke for the button
 */
    private createHTMLButtonElement(
        buttonLabel: string,
        buttonId: string,
        buttonValue: string | null,
        onClickHandler: (event?: any) => void
    ): HTMLButtonElement {
        let button: HTMLButtonElement = document.createElement("button");
        button.innerHTML = buttonLabel;
        if (buttonValue) {
            button.setAttribute("buttonvalue", buttonValue);
        }
        button.id = buttonId;
        button.classList.add("QuickPick_ButtonClass");
        button.addEventListener("click", onClickHandler);
        return button;
    }

}