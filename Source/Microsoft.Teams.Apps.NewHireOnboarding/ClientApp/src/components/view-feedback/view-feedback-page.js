"use strict";
// <copyright file="view-feedback-page.tsx" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>
var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
        return extendStatics(d, b);
    }
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
Object.defineProperty(exports, "__esModule", { value: true });
var React = require("react");
var react_northstar_1 = require("@fluentui/react-northstar");
var view_feedback_api_1 = require("../../api/view-feedback-api");
var react_i18next_1 = require("react-i18next");
var react_spreadsheet_1 = require("react-spreadsheet");
var download_feedback_page_1 = require("../view-feedback/download-feedback-page");
var resources_1 = require("../../constants/resources");
require("bootstrap/dist/css/bootstrap.min.css");
require("../../styles/feedback.css");
var feedbackExcelData = [
    [{ value: "" }, { value: "" }, { value: "" }],
];
var FeedbackPage = /** @class */ (function (_super) {
    __extends(FeedbackPage, _super);
    function FeedbackPage(props) {
        var _this = _super.call(this, props) || this;
        /**
        * get screen width real time.
        */
        _this.update = function () {
            _this.setState({
                screenWidth: window.innerWidth
            });
        };
        /**
        * Fetch share feedback data.
        */
        _this.getFeedbackData = function (batchId) { return __awaiter(_this, void 0, void 0, function () {
            var response;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, view_feedback_api_1.getFeedbackData(batchId)];
                    case 1:
                        response = _a.sent();
                        if (response.status === 200 && response.data) {
                            this.setState({
                                feedbackDetails: response.data
                            });
                        }
                        this.setState({
                            isLoading: false
                        });
                        return [2 /*return*/];
                }
            });
        }); };
        /**
        *Changes dialog open state to show and hide dialog.
        *@param isOpen Boolean indication whether to show dialog
        */
        _this.changeDialogOpenState = function (isOpen) {
            _this.setState({ DownloadDialogOpen: isOpen });
        };
        /**
        *Changes dialog open state to show and hide dialog.
        *@param isOpen Boolean indication whether to show dialog
        */
        _this.closeDialog = function (isOpen) {
            _this.setState({ DownloadDialogOpen: isOpen });
        };
        /**
        *Close the dialog and pass back card properties to parent component.
        */
        _this.onSubmitClick = function () { return __awaiter(_this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        this.batchId = this.state.selectedMonth.substring(0, 3) + "_" + this.state.selectedYear;
                        return [4 /*yield*/, this.setState({ isSubmitClicked: true })];
                    case 1:
                        _a.sent();
                        return [4 /*yield*/, this.getFeedbackData(this.batchId)];
                    case 2:
                        _a.sent();
                        return [4 /*yield*/, this.setState({ isSubmitClicked: false })];
                    case 3:
                        _a.sent();
                        return [2 /*return*/];
                }
            });
        }); };
        _this.localize = _this.props.t;
        window.addEventListener("resize", _this.update);
        _this.batchId = "";
        _this.monthList = _this.localize("months").split(",").map(function (item) { return item.trim(); });
        ;
        _this.state = {
            isLoading: true,
            screenWidth: 0,
            feedbackDetails: [],
            DownloadDialogOpen: false,
            selectedMonth: _this.monthList[new Date().getUTCMonth()],
            selectedYear: resources_1.default.defaultSelectedYear,
            isSubmitClicked: false
        };
        return _this;
    }
    /**
    * Used to initialize Microsoft Teams sdk.
    */
    FeedbackPage.prototype.componentDidMount = function () {
        this.batchId = this.state.selectedMonth.toString().substring(0, 3) + "_" + this.state.selectedYear;
        this.setState({ isLoading: true });
        this.getFeedbackData(this.batchId);
        this.update();
    };
    /**
     * Render feedbacks data.
    */
    FeedbackPage.prototype.renderFeedbacks = function () {
        var _this = this;
        var onMonthChange = {
            onAdd: function (item) {
                _this.setState({
                    selectedMonth: item
                });
                return "";
            }
        };
        var onYearChange = {
            onAdd: function (item) {
                _this.setState({
                    selectedYear: item
                });
                return "";
            }
        };
        if (this.state.feedbackDetails) {
            feedbackExcelData = [
                [{ value: this.localize("columnHeaderMonthText") }, { value: this.localize("columnHeaderNewHireNameText") }, { value: this.localize("columnHeaderFeedbackText") }],
            ];
            this.state.feedbackDetails.forEach(function (feedback) {
                feedbackExcelData.push([{ value: feedback.submittedOn }, { value: feedback.newHireName }, { value: feedback.feedback }]);
            });
        }
        if (this.state.isLoading) {
            return (React.createElement("div", { className: "container-div" },
                React.createElement("div", { className: "container-subdiv" },
                    React.createElement("div", { className: "loader" },
                        React.createElement(react_northstar_1.Loader, null)))));
        }
        else {
            return (React.createElement("div", null,
                React.createElement("div", { className: "feedback-filter-section" },
                    React.createElement(react_northstar_1.Flex, { gap: "gap.smaller", className: "input-fields-margin-between-add-post dateContainer" },
                        React.createElement(react_northstar_1.Flex.Item, { size: "size.quarter" },
                            React.createElement(react_northstar_1.Text, { className: "form-label", content: this.localize("monthLabelText") })),
                        React.createElement(react_northstar_1.Flex.Item, { size: "size.quarter" },
                            React.createElement(react_northstar_1.Text, { className: "form-label", content: this.localize("yearLabelText") }))),
                    React.createElement(react_northstar_1.Flex, { gap: "gap.smaller", className: "input-label-space-between dateContainer" },
                        React.createElement(react_northstar_1.Flex.Item, { size: "size.quarter" },
                            React.createElement(react_northstar_1.Dropdown, { fluid: true, items: this.monthList, value: this.state.selectedMonth, getA11ySelectionMessage: onMonthChange, "data-testid": "monthListTestId" })),
                        React.createElement(react_northstar_1.Flex.Item, { size: "size.quarter" },
                            React.createElement(react_northstar_1.Dropdown, { fluid: true, items: resources_1.default.yearList, value: this.state.selectedYear, getA11ySelectionMessage: onYearChange, "data-testid": "yearListTestId" })),
                        React.createElement(react_northstar_1.Flex.Item, { size: "size.quarter" },
                            React.createElement(react_northstar_1.Button, { className: "dropdown-button", content: this.localize("submitButtonText"), primary: true, loading: this.state.isSubmitClicked, disabled: this.state.isLoading, onClick: this.onSubmitClick })),
                        React.createElement(react_northstar_1.Flex.Item, { size: "size.quarter" },
                            React.createElement(react_northstar_1.Dialog, { className: "dialog-container-view-feedback", content: React.createElement(download_feedback_page_1.default, { batchId: this.batchId, closeDialog: this.closeDialog }), open: this.state.DownloadDialogOpen, onOpen: function () { return _this.setState({ DownloadDialogOpen: false }); }, trigger: React.createElement(react_northstar_1.Button, { className: "dropdown-button", content: this.localize("downloadFeedbackButtonText"), primary: true, onClick: function () { return _this.changeDialogOpenState(true); } }) })))),
                React.createElement("div", { className: "feedback-excel-section" },
                    React.createElement(react_spreadsheet_1.default, { data: feedbackExcelData }))));
        }
    };
    /**
    * Render feedbacks data.
   */
    FeedbackPage.prototype.renderNoFeedbacksFound = function () {
        if (!this.state.feedbackDetails.length) {
            return (React.createElement("div", null,
                React.createElement(react_northstar_1.Text, { className: "feedback-message" },
                    "  ",
                    this.localize("noFeedbackDataFoundMessage"),
                    " ")));
        }
    };
    /**
   * Renders the component.
   */
    FeedbackPage.prototype.render = function () {
        return (React.createElement("div", { className: "container-div" },
            React.createElement("div", { className: "container-subdiv" },
                React.createElement("div", null, this.renderFeedbacks()),
                React.createElement("div", null, this.renderNoFeedbacksFound()))));
    };
    return FeedbackPage;
}(React.Component));
exports.default = react_i18next_1.withTranslation()(FeedbackPage);
//# sourceMappingURL=view-feedback-page.js.map