import React, { useState, useEffect } from 'react';
import * as SDK from 'azure-devops-extension-sdk';
import { Checkbox } from 'azure-devops-ui/Checkbox';
import { TextField } from 'azure-devops-ui/TextField';
import { IExtensionDataService, CommonServiceIds } from 'azure-devops-extension-api';
import { IWorkItemFormService, WorkItemTrackingServiceIds } from "azure-devops-extension-api/WorkItemTracking";
import { Deliverable, showRootComponent, decodeHtmlEntities } from '../Common/Common';
import { WorkItemPicker } from "../Common/WorkItemPicker";

interface Question {
    id: string;
    text: string;
    linkedDeliverables: string[];
}

interface DeliverableDetail {
    value: any;
}

interface Version {
    description: string;
    questions: Question[];
}

const DeliverablesComponent: React.FC = () => {
    const [questions, setQuestions] = useState<Question[]>([]);
    const [selectedQuestions, setSelectedQuestions] = useState<Set<string>>(new Set());
    const [deliverables, setDeliverables] = useState<Deliverable[]>([]);
    const [deliverableDetails, setDeliverableDetails] = useState<{ [deliverableId: string]: DeliverableDetail }>({});
    const [workItemFormService, setWorkItemFormService] = useState<IWorkItemFormService | null>(null);
    const [currentVersionDescription, setCurrentVersionDescription] = useState<string>('');

    useEffect(() => {
        initializeSDK();
    }, []);

    const initializeSDK = async () => {
        await SDK.init();

        try {
            const wifService = await SDK.getService<IWorkItemFormService>(WorkItemTrackingServiceIds.WorkItemFormService);
            setWorkItemFormService(wifService);

            const extDataService = await SDK.getService<IExtensionDataService>(CommonServiceIds.ExtensionDataService);
            const dataManager = await extDataService.getExtensionDataManager(SDK.getExtensionContext().id, await SDK.getAccessToken());

            const storedDeliverables = await dataManager.getValue<Deliverable[]>('deliverables', { scopeType: 'Default' }) || [];
            setDeliverables(storedDeliverables);

            const questionVersions: Version[] = await dataManager.getValue('questionaryVersions', { scopeType: 'Default' }) || [];
            if (questionVersions.length > 0) {
                const latestVersion = questionVersions[questionVersions.length - 1];
                setQuestions(latestVersion.questions);
                setCurrentVersionDescription(latestVersion.description);
            }

            const fieldValue = await wifService.getFieldValue('Custom.AnswersField', { returnOriginalValue: true }) as string;
            if (fieldValue) {
                const decodedFieldValue = decodeHtmlEntities(fieldValue);
                const parsedData = JSON.parse(decodedFieldValue);
                if (parsedData.deliverables) {
                    setDeliverableDetails(parsedData.deliverables);
                }
                if (parsedData.selectedQuestions) {
                    setSelectedQuestions(new Set(parsedData.selectedQuestions));
                }
            }
        } catch (error) {
            console.error("SDK Initialization Error: ", error);
        }
    };

    const handleQuestionSelect = (questionId: string, checked: boolean) => {
        const updatedSelections = new Set(selectedQuestions);
        if (checked) {
            updatedSelections.add(questionId);
        } else {
            updatedSelections.delete(questionId);
        }
        setSelectedQuestions(updatedSelections);
    };

    const handleDeliverableChange = (deliverableId: string, newValue: any) => {
        setDeliverableDetails((prev) => {
            const newDetails = { ...prev, [deliverableId]: { value: newValue } };
            if (workItemFormService) {
                workItemFormService.setFieldValue('Custom.AnswersField', JSON.stringify({ deliverables: newDetails, selectedQuestions: Array.from(selectedQuestions) }));
            }
            return newDetails;
        });
    };

    const renderDeliverableField = (deliverable: Deliverable, detail: DeliverableDetail) => {
        switch (deliverable.type) {
            case 'url':
                return (
                    <TextField
                        value={detail.value || ''}
                        onChange={(_, value) => handleDeliverableChange(deliverable.id, value || '')}
                        placeholder="Enter URL"
                    />
                );
            case 'boolean':
                return (
                    <Checkbox
                        checked={!!detail.value}
                        onChange={(_, checked) => handleDeliverableChange(deliverable.id, checked)}
                        label="Completed"
                    />
                );
            case 'workItem':
                return (
                    <WorkItemPicker
                        onWorkItemSelected={(workItemId) => handleDeliverableChange(deliverable.id, workItemId.toString())}
                    />
                );
            default:
                return null;
        }
    };

    const getUniqueDeliverablesForSelectedQuestions = () => {
        const uniqueDeliverableIds = new Set<string>();
        selectedQuestions.forEach(questionId => {
            const question = questions.find(q => q.id === questionId);
            if (question) {
                question.linkedDeliverables.forEach(deliverableId => uniqueDeliverableIds.add(deliverableId));
            }
        });
        return Array.from(uniqueDeliverableIds)
            .map(deliverableId => deliverables.find(d => d.id === deliverableId))
            .filter((deliverable): deliverable is Deliverable => Boolean(deliverable));
    };

    const uniqueDeliverables = getUniqueDeliverablesForSelectedQuestions();

    // Calculate weights for deliverables
    const weights = uniqueDeliverables.map((_, index) => Math.pow(2, index));

    // Calculate totalWeight
    const totalWeight = weights.reduce((acc, weight) => acc + weight, 0);

    // Calculate uniqueResult (sum of weights for answered deliverables)
    const uniqueResult = uniqueDeliverables.reduce((acc, deliverable, index) => {
        const detail = deliverableDetails[deliverable.id];
        return detail && detail.value ? acc + weights[index] : acc;
    }, 0);

    return (
        <div>
            <h3>Version Description</h3>
            <p>{currentVersionDescription}</p>

            {questions.map((question) => (
                <div key={question.id}>
                    <Checkbox
                        label={question.text}
                        checked={selectedQuestions.has(question.id)}
                        onChange={(_, checked) => handleQuestionSelect(question.id, checked)}
                    />
                </div>
            ))}
            
            <h3>Unique Deliverables</h3>
            <div>
                {uniqueDeliverables.map((deliverable) => (
                    <div key={deliverable.id} style={{ marginBottom: '10px' }}>
                        <div>{deliverable.label} ({deliverable.type})</div>
                        {renderDeliverableField(deliverable, deliverableDetails[deliverable.id] || { value: '' })}
                    </div>
                ))}
            </div>

            <div>
                <p>Total Weight: {totalWeight}</p>
                <p>Unique Result: {uniqueResult}</p>
            </div>
        </div>
    );
};

showRootComponent(<DeliverablesComponent />, 'deliverablesComponent-root');
export default DeliverablesComponent;