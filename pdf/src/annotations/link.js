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
        AscPDF.CAnnotationBase.call(this, sName, AscPDF.ANNOTATIONS_TYPES.Line, aRect, oDoc);
        
        AscPDF.initShape(this);
        this.spPr.setGeometry(AscFormat.CreateGeometry("rect"));

        // states
        this._pressed = false;
        this._hovered = false;
    };
    
    CAnnotationLink.prototype.constructor = CAnnotationLink;
    AscFormat.InitClass(CAnnotationLink, AscPDF.CPdfShape, AscDFH.historyitem_type_Pdf_Annot_Link);
    Object.assign(CAnnotationLink.prototype, AscPDF.CAnnotationBase.prototype);

    CAnnotationLink.prototype.Draw = function(oGraphicsPDF, oGraphicsWord) {
        if (this.IsHidden() && !Asc.editor.IsEditFieldsMode())
            return;

        this.DrawBackground(oGraphicsPDF);
        this.DrawBorders(oGraphicsPDF);

        if (true == this.IsChecked())
            this.DrawCheckedSymbol(oGraphicsPDF);

        this.DrawLocks(oGraphicsPDF);
        this.DrawEdit(oGraphicsWord);
    };
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
        let oDoc = this.GetDocument();

        if (oDoc.IsEditFieldsMode()) {
            this.editShape.onMouseDown(x, y, e);
            return;
        }

        this.DrawPressed();
        this.AddActionsToQueue(AscPDF.FORMS_TRIGGERS_TYPES.MouseDown);
    };
    CAnnotationLink.prototype.onMouseEnter = function() {
        this.AddActionsToQueue(AscPDF.FORMS_TRIGGERS_TYPES.MouseEnter);
        this.SetHovered(true);
    };
    CAnnotationLink.prototype.onMouseExit = function() {
        this.AddActionsToQueue(AscPDF.FORMS_TRIGGERS_TYPES.MouseExit);
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
        let oDoc = this.GetDocument();

        let oThis = this;
        let bCommit = false;
        if (oThis.IsChecked()) {
            if (oThis.IsNoToggleToOff() == false) {
                oThis.SetChecked(false);
                bCommit = true;
            }
        }
        else {
            oThis.SetChecked(true);
            bCommit = true;
        }
        
        this.DrawUnpressed();
    };
    /**
	 * The value application logic for all fields with the same name has been changed for this field type.
     * The method was left for compatibility.
	 * @memberof CRadioButtonField
	 * @typeofeditors ["PDF"]
	 */
    CAnnotationLink.prototype.Commit = function() {
        this.SetNeedCommit(false);

        let oParent = this.GetParent();
        let aOpt    = oParent ? oParent.GetOptions() : undefined;
        let aKids   = oParent ? oParent.GetKids() : undefined;
        if (this.IsChecked()) {
            if (aOpt && aKids) {
                if (this.GetType() == AscPDF.FIELD_TYPES.radiobutton && this.IsRadiosInUnison() || this.GetType() == AscPDF.FIELD_TYPES.checkbox) {
                    this.SetParentValue(aOpt.indexOf(this.GetExportValue()));
                }
                else {
                    this.SetParentValue(String(aKids.indexOf(this)));
                }
            }
            else {
                this.SetParentValue(this.GetExportValue());
            }
        }
        else {
            this.SetParentValue("Off");
        }

        this.Commit2();
    };
    CAnnotationLink.prototype.SetNoToggleToOff = function(bValue) {
        let oParent = this.GetParent();
        if (oParent && oParent.IsAllKidsWidgets()) {
            return oParent.SetNoToggleToOff(bValue);
        }

        if (this._noToggleToOff === bValue) {
            return true;
        }

        AscCommon.History.Add(new CChangesPDFCheckboxNoToggleToOff(this, this._noToggleToOff, bValue));

        this._noToggleToOff = bValue;
        this.SetWasChanged(true);

        return true;
    };
    CAnnotationLink.prototype.IsNoToggleToOff = function(bInherit) {
        let oParent = this.GetParent();
        if (bInherit !== false && oParent && oParent.IsAllKidsWidgets())
            return oParent.IsNoToggleToOff();

        return this._noToggleToOff;
    };
    CAnnotationLink.prototype.SetOptions = function(aOpt) {
        let oParent = this.GetParent();
        if (oParent && oParent.IsAllKidsWidgets()) {
            oParent.SetOptions(aOpt);
        }
        
        let hasOptions = !!this._options;
        
        AscCommon.History.Add(new CChangesPDFCheckOptions(this, this._options, aOpt));

        if (this._options == aOpt) {
            return true;
        }
        
        this._options = aOpt;

        let aAllWidgets = this.GetAllWidgets();
        aAllWidgets.forEach(function(widget) {
            widget.SetExportValue(undefined, true);
        });

        let sDefValue = this.GetDefaultValue();
        let sCurExpValue;

        if (sDefValue) {
            if (!hasOptions) {
                sCurExpValue = this.GetDefaultValue();
            }
            else {
                sCurExpValue = aOpt[sDefValue];
            }

            this.SetDefaultValue(String(aOpt.indexOf(sCurExpValue)));
        }

        return true;
    };
    CAnnotationLink.prototype.GetOptions = function(bInherit) {
        let oParent = this.GetParent();
        if (bInherit !== false && oParent && oParent.IsAllKidsWidgets())
            return oParent.GetOptions();

        return this._options;
    };
    CAnnotationLink.prototype.GetOptionsIndex = function() {
        let oParent = this.GetParent();
        let aOptions = oParent ? oParent.GetOptions() : null;
        if (aOptions) {
            let aKids = oParent.GetKids();
            return aKids.indexOf(this);
        }

        return -1;
    };
    CAnnotationLink.prototype.AddKid = function(oField) {
        let aOptions = this.GetOptions();
        let aNewOptions = aOptions ? aOptions.slice() : null;
        if (aNewOptions) {
            aNewOptions.push(oField.GetExportValue())
        }
        
        AscCommon.History.Add(new CChangesPDFFormKidsContent(this, this._kids.length, [oField], true))

        this._kids.push(oField);
        oField._parent = this;

        if (false == Asc.editor.getDocumentRenderer().IsOpenFormsInProgress) {
            if (oField.IsWidget()) {
                oField.SyncValue();
            }

            if (!aOptions) {
                aOptions = [];

                let bSetOptions = false;

                this._kids.forEach(function(widget) {
                    let sExportValue = widget.GetExportValue();
                    if (aOptions.includes(sExportValue)) {
                        bSetOptions = true;
                    }

                    aOptions.push(sExportValue);
                });

                if (bSetOptions) {
                    aNewOptions = aOptions;
                }
            }
        }
        
        if (aNewOptions) {
            this.SetOptions(aNewOptions);
        }
    };
    CAnnotationLink.prototype.RemoveKid = function(oField) {
        let nIndex = this._kids.indexOf(oField);

        let aOptions = this.GetOptions();
        let aNewOptions = aOptions ? aOptions.slice() : null;
        let sExportValue;
        if (aNewOptions) {
            sExportValue = aNewOptions[nIndex];
            aNewOptions.splice(nIndex, 1);
            this.SetOptions(aNewOptions);
        }
        
        if (nIndex != -1) {
            this._kids.splice(nIndex, 1);
            AscCommon.History.Add(new CChangesPDFFormKidsContent(this, nIndex, [oField], false))
            oField._parent = null;

            if (aNewOptions) {
                oField.SetExportValue(sExportValue);
            }

            return true;
        }

        return false;
    };
    CAnnotationLink.prototype.SetExportValue = function(sValue) {
        let oParent = this.GetParent();
    
        if (oParent && sValue !== undefined) {
            let aWidgets        = oParent.GetAllWidgets();
            let nIndex          = aWidgets.indexOf(this);
            let aExpValues      = aWidgets.map(function(w) { return w.GetExportValue() });
            let aCurOptions     = oParent.GetOptions();

            const newValues = aExpValues.slice();
            newValues[nIndex] = sValue;
    
            if (aExpValues.includes(sValue) || aCurOptions) {
                oParent.SetOptions(newValues);
                return true;
            }
        }
    
        if (this._exportValue == sValue) {
            return false;
        }

        AscCommon.History.Add(new CChangesPDFCheckboxExpValue(this, this._exportValue, sValue));
        this._exportValue = sValue;
        this.SetWasChanged(true);
    };
    CAnnotationLink.prototype.GetExportValue = function(bInherit) {
        if (bInherit !== false) {
            let oParent = this.GetParent();
            let aParentOpt = oParent ? oParent.GetOptions() : null;

            if (aParentOpt) {
                return aParentOpt[oParent.GetKids().indexOf(this)];
            }
        }

        return this._exportValue;
    };
    /**
     * Sets the checkbox style
     * @memberof CAnnotationLink
     * @param {number} nType - checkbox style type (CHECKBOX_STYLES)
     * @typeofeditors ["PDF"]
     */
    CAnnotationLink.prototype.SetStyle = function(nType) {
        AscCommon.History.Add(new CChangesPDFCheckboxStyle(this, this._chStyle, nType));

        this._chStyle = nType;
        this.SetWasChanged(true);
        this.AddToRedraw(true);
    };
    CAnnotationLink.prototype.GetStyle = function() {
        return this._chStyle;
    };
    CAnnotationLink.prototype.SetValue = function(value) {
        let oParent     = this.GetParent();
        let aParentOpt  = oParent ? oParent.GetOptions() : undefined;

        let sExportValue;
        if (aParentOpt && aParentOpt[value]) {
            sExportValue = aParentOpt[value];
        }
        else {
            sExportValue = value;
        }

        if (this.GetExportValue() == sExportValue)
            this.SetChecked(true);
        else
            this.SetChecked(false);
        
        if (editor.getDocumentRenderer().IsOpenFormsInProgress && this.GetParent() == null)
            this.SetParentValue(value);
    };
    CAnnotationLink.prototype.private_SetValue = CAnnotationLink.prototype.SetValue;
    CAnnotationLink.prototype.GetValue = function() {
        return this.IsChecked() ? this.GetExportValue() : "Off";
    };
    CAnnotationLink.prototype.SetDrawFromStream = function() {
    };
    
    /**
     * Set checked to this field (not for all with the same name).
     * @memberof CAnnotationLink
     * @typeofeditors ["PDF"]
     */
    CAnnotationLink.prototype.SetChecked = function(bChecked) {
        if (bChecked == this.IsChecked())
            return;

        this.SetWasChanged(true);
        this.AddToRedraw();

        if (bChecked) {
            AscCommon.History.Add(new CChangesPDFFormValue(this, this.GetValue(), this.GetExportValue()));
            this._checked = true;
        }
        else {
            AscCommon.History.Add(new CChangesPDFFormValue(this, this.GetValue(), "Off"));
            this._checked = false;
        }
    };
    /**
	 * Synchronizes this field with fields with the same name.
	 * @memberof CCheckBoxField
	 * @typeofeditors ["PDF"]
	 */
    CAnnotationLink.prototype.SyncValue = function() {
        if (this.GetExportValue() == this.GetParentValue()) {
            this.SetChecked(true);
            this.AddToRedraw();
        }
        else {
            this.SetChecked(false);
            this.AddToRedraw();
        }
    };
    CAnnotationLink.prototype.DrainLogicFrom = function(oFieldToInherit, bClearFrom) {
        AscPDF.CBaseField.prototype.DrainLogicFrom.call(this, oFieldToInherit, bClearFrom);

        this.SetNoToggleToOff(oFieldToInherit.IsNoToggleToOff());
        if (this.GetType() == AscPDF.FIELD_TYPES.radiobutton) {
            this.SetRadiosInUnison(oFieldToInherit.IsRadiosInUnison());
        }

        if (bClearFrom !== false) {
            oFieldToInherit.SetNoToggleToOff(false);

            if (this.GetType() == AscPDF.FIELD_TYPES.radiobutton) {
                oFieldToInherit.SetRadiosInUnison(false);
            }
        }
    };
    CAnnotationLink.prototype.DrainViewPropsFrom = function(oField) {
        AscPDF.CBaseField.prototype.DrainViewPropsFrom.call(this, oField);

        this.SetStyle(oField.GetStyle());
    };
    CAnnotationLink.prototype.WriteToBinary = function(memory) {
        memory.WriteByte(AscCommon.CommandType.ctAnnotField);

        // длина комманд
        let nStartPos = memory.GetCurPosition();
        memory.Skip(4);

        this.WriteToBinaryBase(memory);
        this.WriteToBinaryBase2(memory);

        // checked
        let isChecked = this.IsChecked();
        // не пишем значение, если есть родитель с такими же видджет полями,
        // т.к. значение будет хранить родитель
        let oParent = this.GetParent();
        if (oParent == null || oParent.IsAllKidsWidgets() == false) {
            memory.fieldDataFlags |= (1 << 9);
            if (isChecked) {
                memory.WriteString("Yes");
            }
            else
                memory.WriteString("Off");
        }
        
        // check symbol
        memory.WriteByte(this.GetStyle());

        let sExportValue = this.GetExportValue(memory.isCopyPaste);
        if (sExportValue != null) {
            memory.fieldDataFlags |= (1 << 14);
            memory.WriteString(sExportValue);
        }

        if (this.IsNoToggleToOff(memory.isCopyPaste)) {
            memory.widgetFlags |= (1 << 14);
        }

        if (this.GetType() == AscPDF.FIELD_TYPES.radiobutton) {
            if (this.IsRadiosInUnison(memory.isCopyPaste)) {
                memory.widgetFlags |= (1 << 25);
            }
        }
        let nEndPos = memory.GetCurPosition();

        // запись флагов
        memory.Seek(memory.posForWidgetFlags);
        memory.WriteLong(memory.widgetFlags);
        memory.Seek(memory.posForFieldDataFlags);
        memory.WriteLong(memory.fieldDataFlags);

        // запись длины комманд
        memory.Seek(nStartPos);
        memory.WriteLong(nEndPos - nStartPos);
        memory.Seek(nEndPos);

        this.CheckWidgetFlags(memory);
    };
    if (!window["AscPDF"])
	    window["AscPDF"] = {};
    
    let CHECK_SVG = "<svg width='20' height='20' viewBox='0 0 20 20' fill='none' xmlns='http://www.w3.org/2000/svg'>\
    <path d='M5.2381 8.8L4 11.8L7.71429 16C12.0476 9.4 13.2857 8.2 17 4C14.5238 4 9.77778 8.8 7.71429 11.8L5.2381 8.8Z' fill='black'/>\
    </svg>";

    function toBase64(str) {
		return window.btoa(unescape(encodeURIComponent(str)));
	}
	
	function getSvgImage(svg) {
		let image = new Image();
		if (!AscCommon.AscBrowser.isIE || AscCommon.AscBrowser.isIeEdge) {
			image.src = "data:image/svg+xml;utf8," + encodeURIComponent(svg);
		}
		else {
			image.src = "data:image/svg+xml;base64," + toBase64(svg);
			image.onload = function() {
				// Почему-то IE не определяет размеры сам
				this.width = 20;
				this.height = 20;
			};
		}
		
		return image;
	}

    const CHECKED_ICON = getSvgImage(CHECK_SVG);

	window["AscPDF"].CAnnotationLink = CAnnotationLink;
})();

