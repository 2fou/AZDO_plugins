import * as React from 'react';
import * as SDK from 'azure-devops-extension-sdk';
import { useState, useEffect } from 'react';
import { IWorkItemFormService, WorkItemTrackingServiceIds } from 'azure-devops-extension-api/WorkItemTracking';
import { decodeHtmlEntities, showRootComponent, Version, AnswerData } from '../Common/Common';
import { CommonServiceIds, IExtensionDataService } from 'azure-devops-extension-api';

interface FieldChangedEventArgs {
    changedFields: Record<string, any>;
}

// Function to get the Work Item Form Service from the SDK
async function getWorkItemFormService() {
    return SDK.getService<IWorkItemFormService>(WorkItemTrackingServiceIds.WorkItemFormService);
}

const ProgressIndicator: React.FC = () => {
    // State variables for loading state, error messages, progress percentage, version data, and answer data
    const [loading, setLoading] = useState<boolean>(true);
    const [error, setError] = useState<string | null>(null);
    const [overallProgress, setOverallProgress] = useState<number>(0);
    const [versionData, setVersionData] = useState<Version | null>(null);
    const [answerData, setAnswerData] = useState<AnswerData | null>(null);

    // Function to load data from a custom field in the work item form
    const loadWidgetData = async (workItemFormService: IWorkItemFormService) => {
        try {
            const fieldValue = await workItemFormService.getFieldValue('Custom.AnswersField', { returnOriginalValue: true });
            handleFieldValueChange(fieldValue);
        } catch (error) {
            console.error("Error loading field value:", error);
            setError("Failed to load field data.");
        }
    };

    // Effect to initialize the SDK on component mount
    useEffect(() => {
        const initializeSDK = async () => {
            try {
                await SDK.init(); // Initialize the SDK
                await SDK.ready(); // Wait for the SDK to be ready

                // Register event handlers for work item events: loaded, saved, and field changed
                SDK.register(SDK.getContributionId(), () => ({
                    onLoaded: async () => {
                        console.log("Work item loaded");
                        setLoading(true);
                        const workItemFormService = await getWorkItemFormService();
                        await loadWidgetData(workItemFormService);
                        setLoading(false);
                    },
                    onSaved: async () => {
                        console.log("Work item saved");
                        setLoading(true);
                        const workItemFormService = await getWorkItemFormService();
                        await loadWidgetData(workItemFormService);
                        setLoading(false);
                    },
                    onFieldChanged: async (args: FieldChangedEventArgs) => {
                        if (args.changedFields['Custom.AnswersField'] !== undefined) {
                            console.log("Field changed:", args.changedFields['Custom.AnswersField']);
                            const workItemFormService = await getWorkItemFormService();
                            await loadWidgetData(workItemFormService);
                        }
                    },
                }));

                setLoading(false);
            } catch (err) {
                console.error("Error during SDK initialization:", err);
                setError("Failed to initialize component.");
                setLoading(false);
            }
        };

        initializeSDK();
    }, []);

    // Function to handle changes in the custom field value
    const handleFieldValueChange = async (fieldValue: any) => {
        try {
            if (fieldValue) {
                const decodedValue = decodeHtmlEntities(fieldValue as string); // Decode HTML entities
                const data: AnswerData = JSON.parse(decodedValue); // Parse the JSON data
                setAnswerData(data);
                await fetchVersionData(data.version); // Fetch version data based on the parsed answer data
            } else {
                setError("No field value found.");
            }
        } catch (parseError) {
            console.error("Error parsing answers:", parseError);
            setError("Error parsing answers.");
        }
    };

    // Function to fetch version data using the provided version description
    const fetchVersionData = async (versionDescription: string) => {
        try {
            const extDataService = await SDK.getService<IExtensionDataService>(CommonServiceIds.ExtensionDataService);
            const dataManager = await extDataService.getExtensionDataManager(SDK.getExtensionContext().id, await SDK.getAccessToken());
            // Retrieve saved versions
            const questionVersions: Version[] = await dataManager.getValue('questionaryVersions', { scopeType: 'Default' }) || [];
            const selectedVersion = questionVersions.find(v => v.description === versionDescription);

            if (selectedVersion) {
                setVersionData(selectedVersion);
            } else {
                console.warn("No matching version found for the description:", versionDescription);
            }
        } catch (error) {
            console.error("Error fetching version data:", error);
            setError("Error fetching version data.");
        }
    };

    // Effect to calculate the overall progress whenever versionData or answerData changes
    useEffect(() => {
        if (versionData && answerData) {
            const progress = calculateOverallProgress(answerData);
            setOverallProgress(progress);
        }
    }, [versionData, answerData]);

    // Function to calculate overall progress percentage
    const calculateOverallProgress = (answerData: AnswerData): number => {
        if (!versionData) return 0;

        try {
            // Count the number of answered questions
            const answeredQuestions = answerData.uniqueResult.toString(2).split('1').length - 1;
            // Count the total number of questions
            const totalQuestions = answerData.totalWeight.toString(2).split('1').length - 1;
            return totalQuestions > 0 ? (answeredQuestions / totalQuestions) * 100 : 0;
        } catch (error) {
            console.error("Error during progress calculation:", error);
            return 0;
        }
    };

    // Function to determine progress bar color based on progress percentage
    const calculateColor = (progress: number) => `hsl(${120 * (progress / 100)}, 100%, 50%)`;

    // Render loading, error, or the progress indicator UI
    if (loading) return <div>Loading progress...</div>;

    if (error) return <div>Error: {error}</div>;

    return (
        <div>
            {versionData ? (
                <div>
                    <h3>{versionData.description}</h3>
                    <div style={{ width: '100%', backgroundColor: '#e0e0e0', borderRadius: '8px' }}>
                        <div
                            style={{
                                width: `${overallProgress}%`,
                                backgroundColor: calculateColor(overallProgress),
                                height: '24px',
                                borderRadius: '8px'
                            }}
                        />
                    </div>
                    <span>{overallProgress.toFixed(2)}%</span>
                </div>
            ) : (
                <div>No version data available.</div>
            )}
        </div>
    );
};

// Show the main component using the SDK's root component rendering method
showRootComponent(<ProgressIndicator />, 'progress-root');