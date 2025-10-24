/*
 * (c) Copyright Ascensio System SIA 2010-2024
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation. In accordance with
 * Section 7(a) of the GNU AGPL its Section 15 shall be amended to the effect
 * that Ascensio System SIA expressly excludes the warranty of non-infringement
 * of any third-party rights.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR  PURPOSE. For
 * details, see the GNU AGPL at: http://www.gnu.org/licenses/agpl-3.0.html
 *
 * You can contact Ascensio System SIA at 20A-6 Ernesta Birznieka-Upish
 * street, Riga, Latvia, EU, LV-1050.
 *
 * The  interactive user interfaces in modified source and object code versions
 * of the Program must display Appropriate Legal Notices, as required under
 * Section 5 of the GNU AGPL version 3.
 *
 * Pursuant to Section 7(b) of the License you must retain the original Product
 * logo when distributing the program. Pursuant to Section 7(e) we decline to
 * grant you any rights under trademark law for use of our trademarks.
 *
 * All the Product's GUI elements, including illustrations and icon sets, as
 * well as technical writing content are licensed under the terms of the
 * Creative Commons Attribution-ShareAlike 4.0 International. See the License
 * terms at http://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 */

(function(){

    /**
	 * Class representing a link annotation.
	 * @constructor
     * @extends {CAnnotationBase}
	 */
    function CAnnotationLink(sName, aRect, oDoc)
    {
        AscPDF.CPdfShape.call(this);
        AscPDF.CAnnotationBase.call(this, sName, AscPDF.ANNOTATIONS_TYPES.Link, aRect, oDoc);
        
        AscPDF.initShape(this);
        let oGeometry = AscFormat.CreateGeometry("rect");
        oGeometry.preset = undefined;
        this.spPr.setGeometry(oGeometry);

        this._triggers      = new AscPDF.CPdfTriggers();
        this._quads         = [];

        // states
        this._pressed = false;
        this._hovered = false;
    };
    
    CAnnotationLink.prototype.constructor = CAnnotationLink;
    AscFormat.InitClass(CAnnotationLink, AscPDF.CPdfShape, AscDFH.historyitem_type_Pdf_Annot_Link);
    Object.assign(CAnnotationLink.prototype, AscPDF.CAnnotationBase.prototype);

    CAnnotationLink.prototype.IsLink = function() {
        return true;
    };
    CAnnotationLink.prototype.SetQuads = function(aFullQuads) {
        let oThis = this;
        aFullQuads.forEach(function(aQuads) {
            oThis.AddQuads(aQuads);
        });
    };
    CAnnotationLink.prototype.GetQuads = function() {
        return this._quads;
    };
    CAnnotationLink.prototype.AddQuads = function(aQuads) {
        AscCommon.History.Add(new CChangesPDFAnnotQuads(this, this._quads.length, aQuads, true));
        this._quads.push(aQuads);
    };

    CAnnotationLink.prototype.RefillGeometry = function() {};
    CAnnotationLink.prototype.SetPressed = function(bValue) {
        this._pressed = bValue;
        this.AddToRedraw();
    };
    CAnnotationLink.prototype.IsPressed = function() {
        return this._pressed;
    };
    CAnnotationLink.prototype.IsHovered = function() {
        return this._hovered;
    };
    CAnnotationLink.prototype.SetHovered = function(bValue) {
        this._hovered = bValue;
    };

    CAnnotationLink.prototype.onMouseDown = function(x, y, e) {
        if (Asc.editor.canEdit()) {
            AscPDF.CPdfShape.prototype.onMouseDown.call(this, x, y, e);
            return;
        }

        this.DrawPressed();
    };
    CAnnotationLink.prototype.onMouseEnter = function() {
        this.SetHovered(true);
    };
    CAnnotationLink.prototype.onMouseExit = function() {
        this.SetHovered(false);
    };
    CAnnotationLink.prototype.DrawPressed = function() {
        this.SetPressed(true);
        Asc.editor.getDocumentRenderer()._paint();
    };
    CAnnotationLink.prototype.DrawUnpressed = function() {
        this.SetPressed(false);
        Asc.editor.getDocumentRenderer()._paint();
    };
    CAnnotationLink.prototype.onMouseUp = function() {
        this.DrawUnpressed();
        this.AddActionsToQueue(AscPDF.PDF_TRIGGERS_TYPES.MouseUp);
    };
    CAnnotationLink.prototype.SetActions = function(nTriggerType, aActionsInfo) {
        let aActions = [];
        if (aActionsInfo) {
            for (let i = 0; i < aActionsInfo.length; i++) {
                let oAction;
                switch (aActionsInfo[i]["S"]) {
                    case AscPDF.ACTIONS_TYPES.JavaScript:
                        oAction = new AscPDF.CActionRunScript(aActionsInfo[i]["JS"]);
                        aActions.push(oAction);
                        break;
                    case AscPDF.ACTIONS_TYPES.ResetForm:
                        oAction = new AscPDF.CActionReset(aActionsInfo[i]["Fields"], Boolean(aActionsInfo[i]["Flags"]));
                        aActions.push(oAction);
                        break;
                    case AscPDF.ACTIONS_TYPES.URI:
                        oAction = new AscPDF.CActionURI(aActionsInfo[i]["URI"]);
                        aActions.push(oAction);
                        break;
                    case AscPDF.ACTIONS_TYPES.HideShow:
                        oAction = new AscPDF.CActionHideShow(Boolean(aActionsInfo[i]["H"]), aActionsInfo[i]["T"]);
                        aActions.push(oAction);
                        break;
                    case AscPDF.ACTIONS_TYPES.GoTo:
                        let oRect = {
                            top:    aActionsInfo[i]["top"],
                            right:  aActionsInfo[i]["right"],
                            bottom: aActionsInfo[i]["bottom"],
                            left:   aActionsInfo[i]["left"]
                        }
                        if (aActionsInfo[i]["bottom"] != null && aActionsInfo[i]["top"] != null) {
                            oRect.top = aActionsInfo[i]["bottom"];
                            oRect.bottom = aActionsInfo[i]["top"];
                        }
    
                        oAction = new AscPDF.CActionGoTo(aActionsInfo[i]["page"], aActionsInfo[i]["kind"], aActionsInfo[i]["zoom"], oRect);
                        aActions.push(oAction);
                        break;
                    case AscPDF.ACTIONS_TYPES.Named:
                        oAction = new AscPDF.CActionNamed(AscPDF.CActionNamed.GetInternalType(aActionsInfo[i]["N"]));
                        aActions.push(oAction);
                        break;
                }
            }
        }
        
        const oNewTrigger = aActions.length != 0 ? new AscPDF.CPdfTrigger(nTriggerType, aActions) : null;
        if (oNewTrigger) {
            oNewTrigger.SetParentField(this);
        }

        const aCurActionsInfo = this.GetActions(nTriggerType);
        AscCommon.History.Add(new CChangesPDFFormActions(this, aCurActionsInfo, aActionsInfo, nTriggerType));

        switch (nTriggerType) {
            case AscPDF.PDF_TRIGGERS_TYPES.MouseUp:
                this._triggers.MouseUp = oNewTrigger;
                break;
            case AscPDF.PDF_TRIGGERS_TYPES.MouseDown:
                this._triggers.MouseDown = oNewTrigger;
                break;
            case AscPDF.PDF_TRIGGERS_TYPES.MouseEnter:
                this._triggers.MouseEnter = oNewTrigger;
                break;
            case AscPDF.PDF_TRIGGERS_TYPES.MouseExit:
                this._triggers.MouseExit = oNewTrigger;
                break;
            case AscPDF.PDF_TRIGGERS_TYPES.OnFocus:
                this._triggers.OnFocus = oNewTrigger;
                break;
            case AscPDF.PDF_TRIGGERS_TYPES.OnBlur:
                this._triggers.OnBlur = oNewTrigger;
                break;
        }

        return aActions;
    };
    CAnnotationLink.prototype.GetActions = function(nTriggerType) {
        // Get the trigger by type
        let oTrigger = this.GetTrigger(nTriggerType);
        if (!oTrigger || !oTrigger.Actions) {
            return [];
        }
        
        let aActionsInfo = [];
        // Iterate through all actions associated with the trigger
        for (let i = 0; i < oTrigger.Actions.length; i++) {
            let oAction = oTrigger.Actions[i];
            let actionInfo = {};
            
            // Determine the action type and populate the object with information
            switch (oAction.GetType()) {
                case AscPDF.ACTIONS_TYPES.JavaScript:
                    actionInfo["S"] = AscPDF.ACTIONS_TYPES.JavaScript;
                    actionInfo["JS"] = oAction.GetScript();
                    break;
                case AscPDF.ACTIONS_TYPES.ResetForm:
                    actionInfo["S"] = AscPDF.ACTIONS_TYPES.ResetForm;
                    actionInfo["Fields"] = oAction.GetNames();
                    actionInfo["Flags"] = Number(oAction.GetNeedAllExcept());
                    break;
                case AscPDF.ACTIONS_TYPES.URI:
                    actionInfo["S"] = AscPDF.ACTIONS_TYPES.URI;
                    actionInfo["URI"] = oAction.GetURI();
                    break;
                case AscPDF.ACTIONS_TYPES.HideShow:
                    actionInfo["S"] = AscPDF.ACTIONS_TYPES.HideShow;
                    actionInfo["H"] = oAction.GetHidden();
                    actionInfo["T"] = oAction.GetNames();
                    break;
                case AscPDF.ACTIONS_TYPES.GoTo:
                    actionInfo["S"] = AscPDF.ACTIONS_TYPES.GoTo;
                    actionInfo["page"] = oAction.GetPage();
                    actionInfo["kind"] = oAction.GetKind();
                    actionInfo["zoom"] = oAction.GetZoom();
                    let oRect = oAction.GetRect();
                    actionInfo["top"] = oRect.top;
                    actionInfo["right"] = oRect.right;
                    actionInfo["bottom"] = oRect.bottom;
                    actionInfo["left"] = oRect.left;
                    break;
                case AscPDF.ACTIONS_TYPES.Named:
                    actionInfo["S"] = AscPDF.ACTIONS_TYPES.Named;
                    actionInfo["N"] = oAction.GetName();
                    break;
                default:
                    // If the type is not recognized, add handling or skip
                    break;
            }
            
            aActionsInfo.push(actionInfo);
        }
        
        return aActionsInfo;
    };
    CAnnotationLink.prototype.GetTrigger = function(nType) {
        switch (nType) {
            case AscPDF.PDF_TRIGGERS_TYPES.MouseUp:
                return this._triggers.MouseUp;
            case AscPDF.PDF_TRIGGERS_TYPES.MouseDown:
                return this._triggers.MouseDown;
            case AscPDF.PDF_TRIGGERS_TYPES.MouseEnter:
                return this._triggers.MouseEnter;
            case AscPDF.PDF_TRIGGERS_TYPES.MouseExit:
                return this._triggers.MouseExit;
            case AscPDF.PDF_TRIGGERS_TYPES.OnFocus:
                return this._triggers.OnFocus;
            case AscPDF.PDF_TRIGGERS_TYPES.OnBlur:
                return this._triggers.OnBlur;
        }

        return null;
    };
    CAnnotationLink.prototype.GetListActions = function() {
        let aActions = [];

        let oAction = this.GetTrigger(AscPDF.PDF_TRIGGERS_TYPES.MouseUp);
        if (oAction) {
            aActions.push(oAction);
        }
        
        oAction = this.GetTrigger(AscPDF.PDF_TRIGGERS_TYPES.MouseDown);
        if (oAction) {
            aActions.push(oAction);
        }

        oAction = this.GetTrigger(AscPDF.PDF_TRIGGERS_TYPES.MouseEnter);
        if (oAction) {
            aActions.push(oAction);
        }

        oAction = this.GetTrigger(AscPDF.PDF_TRIGGERS_TYPES.MouseExit);
        if (oAction) {
            aActions.push(oAction);
        }

        oAction = this.GetTrigger(AscPDF.PDF_TRIGGERS_TYPES.OnFocus);
        if (oAction) {
            aActions.push(oAction);
        }

        oAction = this.GetTrigger(AscPDF.PDF_TRIGGERS_TYPES.OnBlur);
        if (oAction) {
            aActions.push(oAction);
        }

        return aActions;
    };
    CAnnotationLink.prototype.AddActionsToQueue = function() {
        let oThis           = this;
        let oDoc            = this.GetDocument();
        let oActionsQueue   = oDoc.GetActionsQueue();

        Object.values(arguments).forEach(function(type) {
            let oTrigger = oThis.GetTrigger(type);
        
            if (oTrigger && oTrigger.Actions.length > 0 && false == AscCommon.History.UndoRedoInProgress) {
                oActionsQueue.AddActions(oTrigger.Actions);
            }
        })
        
        if (oActionsQueue.actions.length !== 0) {
            oActionsQueue.Start();
        }
    };
    
    CAnnotationLink.prototype.hitInPath = function(x, y) {
        let invert_transform = this.getInvertTransform();
        if (!invert_transform) {
            return false;
        }
        let x_t = invert_transform.TransformPointX(x, y);
        let y_t = invert_transform.TransformPointY(x, y);
        let oGeometry = this.getGeometry();
        return oGeometry.hitInInnerArea(this.getCanvasContext(), x_t, y_t);
    };

    if (!window["AscPDF"])
	    window["AscPDF"] = {};
    
	window["AscPDF"].CAnnotationLink = CAnnotationLink;
})();

