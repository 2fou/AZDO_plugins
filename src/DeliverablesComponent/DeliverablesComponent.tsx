import React, { useState, useEffect } from 'react';
import * as SDK from 'azure-devops-extension-sdk';
import { Checkbox } from 'azure-devops-ui/Checkbox';
import { TextField } from 'azure-devops-ui/TextField';
import { CommonServiceIds, IExtensionDataService } from 'azure-devops-extension-api';
import { IWorkItemFormService, WorkItemTrackingServiceIds } from "azure-devops-extension-api/WorkItemTracking";
import { Deliverable, showRootComponent, decodeHtmlEntities, Question } from '../Common/Common';
import { WorkItemPicker } from "../Common/WorkItemPicker";
import ConfirmDialog from '../Common/ConfirmDialog';

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

    const [versions, setVersions] = useState<Version[]>([]);
    const [currentVersion, setCurrentVersion] = useState<Version | null>(null);
    const [isConfirmDialogVisible, setConfirmDialogVisible] = useState<boolean>(false);
    const [pendingVersion, setPendingVersion] = useState<Version | null>(null);

    // State for weights and results
    const [weights, setWeights] = useState<number[]>([]);
    const [totalWeight, setTotalWeight] = useState<number>(0);
    const [uniqueResult, setUniqueResult] = useState<number>(0);

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

           const storedDeliverables = (await dataManager.getValue<Deliverable[]>('deliverables', { scopeType: 'Default' }) || []).map(deliverable => ({
    ...deliverable
    
}));
setDeliverables(storedDeliverables);
            setDeliverables(storedDeliverables);

            const questionVersions: Version[] = await dataManager.getValue('questionaryVersions', { scopeType: 'Default' }) || [];
            setVersions(questionVersions);

            const selectedVersionDesc = await dataManager.getValue<string>('selectedVersion', { scopeType: 'Default' });
            const selectedVersion = questionVersions.find(v => v.description === selectedVersionDesc);

            if (selectedVersion) {
                setCurrentVersion(selectedVersion);
                setQuestions(selectedVersion.questions);
            } else if (questionVersions.length > 0) {
                const latestVersion = questionVersions[questionVersions.length - 1];
                setCurrentVersion(latestVersion);
                setQuestions(latestVersion.questions);
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
                if (parsedData.weights) {
                    setWeights(parsedData.weights);
                    setTotalWeight(parsedData.totalWeight);
                    setUniqueResult(parsedData.uniqueResult);
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

        updateAnswersField(updatedSelections);
    };

    const updateAnswersField = (updatedSelections: Set<string>) => {
    const currentVersionDescription = currentVersion ?. description || '';

    if (workItemFormService) {
        workItemFormService.setFieldValue('Custom.AnswersField', JSON.stringify({
            version: currentVersionDescription,
            deliverables: deliverableDetails,
            selectedQuestions: Array.from(updatedSelections),
            weights,
            totalWeight,
            uniqueResult
        }));
    }
};
    const handleDeliverableChange = (deliverableId: string, newValue: any) => {
        setDeliverableDetails((prev) => {
            const newDetails = { ...prev, [deliverableId]: { value: newValue } };
            updateAnswersField(selectedQuestions);
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

    const handleVersionChange = (event: React.ChangeEvent<HTMLSelectElement>) => {
        const selectedVersionDescription = event.target.value;
        const selectedVersion = versions.find((version) => version.description === selectedVersionDescription);
        if (selectedVersion) {
            setPendingVersion(selectedVersion);
            setConfirmDialogVisible(true);
        }
    };

    const confirmVersionChange = async () => {
        if (pendingVersion) {
            setCurrentVersion(pendingVersion);
            setQuestions(pendingVersion.questions);
            setSelectedQuestions(new Set());

            const extDataService = await SDK.getService<IExtensionDataService>(CommonServiceIds.ExtensionDataService);
            const dataManager = await extDataService.getExtensionDataManager(SDK.getExtensionContext().id, await SDK.getAccessToken());
            await dataManager.setValue('selectedVersion', pendingVersion.description, { scopeType: 'Default' });

            setPendingVersion(null);
        }
        setConfirmDialogVisible(false);
    };

    return (
        <div>
            <h3>Select Version</h3>
            <select onChange={handleVersionChange} defaultValue={currentVersion?.description}>
                {versions.map(version => (
                    <option key={version.description} value={version.description}>
                        {version.description}
                    </option>
                ))}
            </select>

            <h3>Version Description</h3>
            <p>{currentVersion?.description}</p>

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
                {getUniqueDeliverablesForSelectedQuestions().map((deliverable) => (
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

            {isConfirmDialogVisible && (
                <ConfirmDialog
                    title="Confirm Version Change"
                    message="Are you sure you want to switch to this version? Unsaved changes will be lost."
                    onCancel={() => setConfirmDialogVisible(false)}
                    onConfirm={confirmVersionChange} visible={isConfirmDialogVisible} />
            )}
        </div>
    );
};

showRootComponent(<DeliverablesComponent />, 'deliverablesComponent-root');
export default DeliverablesComponent;