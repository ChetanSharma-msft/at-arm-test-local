// <copyright file="view-feedback-page.tsx" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

import * as React from "react";
import { Dialog, Loader, Flex, Text, Dropdown, Button } from "@fluentui/react-northstar";
import { getFeedbackData } from "../../api/view-feedback-api";
import { WithTranslation, withTranslation } from "react-i18next";
import { TFunction } from "i18next";
import Spreadsheet from "react-spreadsheet";
import DownloadFeedbackPage from "../view-feedback/download-feedback-page";
import Resources from "../../constants/resources";

import 'bootstrap/dist/css/bootstrap.min.css';
import "../../styles/feedback.css";

let feedbackExcelData = [
    [{ value: "" }, { value: "" }, { value: "" }],];

export interface IFeedbackDetails {
    submittedOn: string,
    feedback: string,
    newHireName: string,
}

interface IFeedbackState {
    isLoading: boolean;
    screenWidth: number;
    feedbackDetails: Array<IFeedbackDetails>,
    DownloadDialogOpen: boolean;
    selectedMonth: string;
    selectedYear: number;
    isSubmitClicked: boolean;
}

class FeedbackPage extends React.Component<WithTranslation, IFeedbackState> {
    localize: TFunction;
    batchId: string;
    monthList: Array<string> | undefined;

    constructor(props: any) {
        super(props);
        this.localize = this.props.t;
        window.addEventListener("resize", this.update);
        this.batchId = "";
        this.monthList = this.localize("months").split(",").map(item => item.trim());;

        this.state = {
            isLoading: true,
            screenWidth: 0,
            feedbackDetails: [],
            DownloadDialogOpen: false,
            selectedMonth: this.monthList[new Date().getUTCMonth()],
            selectedYear: Resources.defaultSelectedYear,
            isSubmitClicked: false
        }
    }

    /**
    * Used to initialize Microsoft Teams sdk.
    */
    componentDidMount() {
        this.batchId = this.state.selectedMonth.toString().substring(0, 3) + "_" + this.state.selectedYear;
        this.setState({ isLoading: true });
        this.getFeedbackData(this.batchId);
        this.update();
    }

    /**
    * get screen width real time.
    */
    update = () => {
        this.setState({
            screenWidth: window.innerWidth
        });
    };

    /**
    * Fetch share feedback data.
    */
    getFeedbackData = async (batchId: string) => {
        let response = await getFeedbackData(batchId);
        if (response.status === 200 && response.data) {
            this.setState(
                {
                    feedbackDetails: response.data
                });
        }

        this.setState({
            isLoading: false
        });
    }

    /**
    *Changes dialog open state to show and hide dialog.
    *@param isOpen Boolean indication whether to show dialog
    */
    changeDialogOpenState = (isOpen: boolean) => {
        this.setState({ DownloadDialogOpen: isOpen })
    }

    /**
    *Changes dialog open state to show and hide dialog.
    *@param isOpen Boolean indication whether to show dialog
    */
    closeDialog = (isOpen: boolean) => {
        this.setState({ DownloadDialogOpen: isOpen })
    }

    /**
	*Close the dialog and pass back card properties to parent component.
	*/
    onSubmitClick = async () => {

        this.batchId = this.state.selectedMonth.substring(0, 3) + "_" + this.state.selectedYear;

        await this.setState({ isSubmitClicked: true });
        await this.getFeedbackData(this.batchId);
        await this.setState({ isSubmitClicked: false });
    }

    /**
     * Render feedbacks data. 
    */
    renderFeedbacks() {

        const onMonthChange = {
            onAdd: item => {
                this.setState({
                    selectedMonth: item
                })

                return "";
            }
        }

        const onYearChange = {
            onAdd: item => {
                this.setState({
                    selectedYear: item
                })

                return "";
            }
        }

        if (this.state.feedbackDetails) {

            feedbackExcelData = [
                [{ value: this.localize("columnHeaderMonthText") }, { value: this.localize("columnHeaderNewHireNameText") }, { value: this.localize("columnHeaderFeedbackText") }],
            ];

            this.state.feedbackDetails.forEach(function (feedback) {
                feedbackExcelData.push([{ value: feedback.submittedOn }, { value: feedback.newHireName }, { value: feedback.feedback }]);
            });
        }

        if (this.state.isLoading) {
            return (
                <div className="container-div">
                    <div className="container-subdiv">
                        <div className="loader">
                            <Loader />
                        </div>
                    </div>
                </div>
            );
        }
        else {
            return (
                <div>
                    <div className="feedback-filter-section">
                        <Flex gap="gap.smaller" className="input-fields-margin-between-add-post dateContainer">
                            <Flex.Item size="size.quarter">
                                <Text className="form-label" content={this.localize("monthLabelText")} />
                            </Flex.Item>
                            <Flex.Item size="size.quarter">
                                <Text className="form-label" content={this.localize("yearLabelText")} />
                            </Flex.Item>
                        </Flex>

                        <Flex gap="gap.smaller" className="input-label-space-between dateContainer">
                            <Flex.Item size="size.quarter">
                                <Dropdown
                                    fluid
                                    items={this.monthList}
                                    value={this.state.selectedMonth}
                                    getA11ySelectionMessage={onMonthChange}
                                />
                            </Flex.Item>
                            <Flex.Item size="size.quarter">
                                <Dropdown
                                    fluid
                                    items={Resources.yearList}
                                    value={this.state.selectedYear}
                                    getA11ySelectionMessage={onYearChange}
                                />
                            </Flex.Item>
                            <Flex.Item size="size.quarter">
                                <Button className="dropdown-button" content={this.localize("submitButtonText")} primary loading={this.state.isSubmitClicked} disabled={this.state.isLoading} onClick={this.onSubmitClick} />
                            </Flex.Item>
                            <Flex.Item size="size.quarter">
                                <Dialog
                                    className="dialog-container-view-feedback"
                                    content={<DownloadFeedbackPage batchId={this.batchId} closeDialog={this.closeDialog} />}
                                    open={this.state.DownloadDialogOpen}
                                    onOpen={() => this.setState({ DownloadDialogOpen: false })}
                                    trigger=
                                    {
                                        <Button className="dropdown-button" disabled={this.state.feedbackDetails.length < 1} content={this.localize("downloadFeedbackButtonText")} primary onClick={() => this.changeDialogOpenState(true)} />
                                    }
                                />
                            </Flex.Item>
                        </Flex>
                    </div>
                    <div className="feedback-excel-section">
                        {this.renderFeedbackSection(feedbackExcelData)}
                    </div>
                </div>
            );
        }
    }

    /**
    * Render feedbacks data. 
   */
    renderFeedbackSection(feedbackExcelData: any) {

        if (!this.state.feedbackDetails.length) {
            return (
                <div>
                    <Text className="feedback-message">  {this.localize("noFeedbackDataFoundMessage")} </Text>
                </div>
            );
        }
        else {
            return (
                <Spreadsheet data={feedbackExcelData} />
            );
        }
    }

    /**
   * Renders the component.
   */
    public render(): JSX.Element {
        return (
            <div className="container-div">
                <div className="container-subdiv">
                    <div>
                        {this.renderFeedbacks()}
                    </div>
                </div>
            </div>
        );
    }
}

export default withTranslation()(FeedbackPage)