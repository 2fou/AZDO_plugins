import * as React from 'react';
import { useState, useEffect } from 'react';
import * as SDK from 'azure-devops-extension-sdk';
import { IWorkItemFormService, WorkItemTrackingServiceIds } from 'azure-devops-extension-api/WorkItemTracking';
import { decodeHtmlEntities, showRootComponent, Version, Question, AnswerData } from '../Common/Common';
import { CommonServiceIds, IExtensionDataService } from 'azure-devops-extension-api';

console.log("Initializing SDK...");
SDK.init();


const ProgressIndicator: React.FC = () => {
    const [loading, setLoading] = useState<boolean>(true);
    const [error, setError] = useState<string | null>(null);
    const [questionsProgress, setQuestionsProgress] = useState<Array<{ questionText: string, progress: number }>>([]);
    const [versionData, setVersionData] = useState<Version | null>(null);
    const [answerData, setAnswerData] = useState<AnswerData | null>(null);




useEffect(() => {
    const initializeSDK = async () => {
        try {
            setLoading(true);


            console.log("SDK Initialized!");
            await SDK.ready();

            console.log("SDK Ready!");

            const workItemFormService = await SDK.getService<IWorkItemFormService>(WorkItemTrackingServiceIds.WorkItemFormService);

            console.log("WorkItemFormService obtained!");

            // Load initial field value
            await loadFieldValue(workItemFormService);

            // Register the field change handler with a unique registeredObjectId
            const registeredObjectId = 'progress-indicator-object';
            
            console.log("Registering object with ID:", registeredObjectId);

            SDK.register(SDK.getContributionId(), {
                registeredObjectId,
                onFieldChanged: async (args: any) => {
                    if (args.changedFields['Custom.AnswersField'] !== undefined) {
                        console.log("Custom.AnswersField changed:", args.changedFields['Custom.AnswersField']);
                        await loadFieldValue(workItemFormService);
                    }
                }
            });

        } catch (initError) {
            console.error("Error initializing component:", initError);
            setError("Failed to initialize component.");
        } finally {
            setLoading(false);
        }
    };

    const loadFieldValue = async (workItemFormService: IWorkItemFormService) => {
        try {
            const fieldValue = await workItemFormService.getFieldValue('Custom.AnswersField', { returnOriginalValue: true });
            handleFieldValueChange(fieldValue);
        } catch (error) {
            console.error("Error loading field value:", error);
            setError("Failed to load field data.");
        }
    };

    initializeSDK();
}, []);
    
    const handleFieldValueChange = async (fieldValue: any) => {
        if (fieldValue) {
            try {
                const decodedValue = decodeHtmlEntities(fieldValue as string);
                console.log("Decoded field value:", decodedValue);
    
                const data: AnswerData = JSON.parse(decodedValue);
                console.log("Parsed AnswerData:", data);
    
                setAnswerData(data);
                await fetchVersionData(data.version);
            } catch (parseError) {
                console.error("Error parsing answers:", parseError);
                setError("Error parsing answers.");
            }
        } else {
            setError("No field value found.");
        }
    };

    useEffect(() => {
        if (versionData && answerData) {
            console.log("Both versionData and answerData are available, proceeding to calculate progress.");

            const progressData = calculateProgress(answerData);
            setQuestionsProgress(progressData);

            console.log("Calculated Progress Data:", progressData);
        } else {
            console.log("Either versionData or answerData is not available. Current values:", { versionData, answerData });
        }
    }, [versionData, answerData]);

    const fetchVersionData = async (versionDescription: string) => {
        console.log("Fetching version data for:", versionDescription);

        const extDataService = await SDK.getService<IExtensionDataService>(CommonServiceIds.ExtensionDataService);
        const dataManager = await extDataService.getExtensionDataManager(SDK.getExtensionContext().id, await SDK.getAccessToken());
        console.log("Obtained Extension DataService Manager.");

        const questionVersions: Version[] = await dataManager.getValue('questionaryVersions', { scopeType: 'Default' }) || [];
        console.log("Fetched question versions:", questionVersions);

        const selectedVersion = questionVersions.find(v => v.description === versionDescription);
        console.log("Selected Version:", selectedVersion);

        if (selectedVersion) {
            setVersionData(selectedVersion);
            console.log("Version data successfully set:", selectedVersion);
        } else {
            console.warn("No matching version found for the description:", versionDescription);
        }
    };

    const calculateProgress = (answerData: AnswerData) => {
        console.log("Calculating progress...");

        if (!versionData) {
            console.log("No version data available.");
            return [];
        }

        try {
            const answeredQuestions = answerData.uniqueResult.toString(2).split('1').length - 1;
            const totalQuestions = answerData.totalWeight.toString(2).split('1').length - 1;
            const progress = totalQuestions > 0 ? (answeredQuestions / totalQuestions) * 100 : 0;

            console.log(`Answered Questions: ${answeredQuestions}, Total Questions: ${totalQuestions}, Progress: ${progress}`);

            return versionData.questions.map((question: Question) => {
                const questionProgress = answerData.selectedQuestions.includes(question.id) ? progress : 0;

                return {
                    questionText: question.text,
                    progress: questionProgress
                };
            });
        } catch (error) {
            console.error("Error during progress calculation:", error);
            return [];
        }
    };

    const calculateColor = (progress: number) => {
        const hue = 120 * (progress / 100);
        return `hsl(${hue}, 100%, 50%)`;
    };

    if (loading) {
        return <div>Loading progress...</div>;
    }

    if (error) {
        return <div>Error: {error}</div>;
    }

    return (
        <div>
            {questionsProgress.map(({ questionText, progress }) => (
                <div key={questionText} style={{ marginBottom: '10px' }}>
                    <span>{questionText}: {progress.toFixed(2)}%</span>
                    <div style={{ width: '100%', backgroundColor: '#e0e0e0', borderRadius: '8px' }}>
                        <div
                            style={{
                                width: `${progress}%`,
                                backgroundColor: calculateColor(progress),
                                height: '24px',
                                borderRadius: '8px'
                            }}
                        />
                    </div>
                </div>
            ))}
        </div>
    );
};

showRootComponent(<ProgressIndicator />, 'progress-root');