import * as React from 'react';
import { useState, useEffect } from 'react';
import * as SDK from 'azure-devops-extension-sdk';
import { IProjectPageService, getClient, CommonServiceIds } from 'azure-devops-extension-api';
import { WorkItemTrackingRestClient, WorkItem } from 'azure-devops-extension-api/WorkItemTracking';
import { AnswerDetail, showRootComponent, decodeHtmlEntities } from '../Common/Common';

interface WorkItemProgress {
    id: number;
    title: string;
    progress: number;
    completed: number;
    total: number;
}

const QueryAndFieldDashboardWidget: React.FC = () => {
    const [loading, setLoading] = useState<boolean>(true);
    const [error, setError] = useState<string | null>(null);
    const [workItemsProgress, setWorkItemsProgress] = useState<WorkItemProgress[]>([]);

    useEffect(() => {
        let isMounted = true; // Track if the component is mounted

        const initializeAndFetchData = async () => {
            try {
                await initializeSDK();
                console.log("preparing to fetch project info...");
                const projectInfo = await fetchProjectInfo();
               
                if (projectInfo) {
                    console.log("preparing to fetch work items...");
                    await fetchAndProcessWorkItems(projectInfo.id);
                }
            } catch (err) {
                console.error("Error:", err);
                if (isMounted) {
                    setError("Failed to load data.");
                }
            } finally {
                if (isMounted) {
                    setLoading(false);
                }
            }
        };

        const initializeSDK = async () => {
            await SDK.init();
            await SDK.ready();
            console.log("SDK initialized and ready...");
        };

        const fetchProjectInfo = () => {
            const projectServicePromise = SDK.getService<IProjectPageService>(CommonServiceIds.ProjectPageService);
            
            return projectServicePromise.then(projectService => {
                console.log("projectService instantiated...");
                
                return projectService.getProject().then(project => {
                    
                    console.log("Project info fetched...");
                    return project; // Return the fetched project information
                });
            }).catch(error => {
                console.error("Error fetching project info:", error);
                throw error; // Optional: rethrow the error for further handling
            });
        };

        const fetchAndProcessWorkItems = async (projectId: string) => {
            const client = getClient(WorkItemTrackingRestClient);
            console.log("WorkItemTrackingRestClient instantiated...");

            const wiqlQuery = {
                query: `SELECT [System.Id], [System.Title], [Custom.AnswersField] FROM WorkItems WHERE [System.TeamProject] = @project AND [Custom.AnswersField] IS NOT NULL ORDER BY [System.Id]`
            };

            const queryResult = await client.queryByWiql(wiqlQuery, projectId);
            const workItemRefs = queryResult.workItems;
            console.log(`Number of work items found: ${workItemRefs.length}`);

            if (isMounted && workItemRefs.length > 0) {
                const workItems = await client.getWorkItems(
                    workItemRefs.map(wi => wi.id),
                    undefined,
                    ['System.Title', 'Custom.AnswersField']
                );

                if (isMounted) {
                    if (!workItems || workItems.length === 0) {
                        console.warn("No work items were returned.");
                    } else {
                        console.log(`Processing ${workItems.length} work items...`);
                        const progressData = workItems.map(workItem => mapWorkItemToProgress(workItem));
                        setWorkItemsProgress(progressData);
                    }
                }
            } else {
                console.warn("No work item references were found.");
            }
        };

        // Call the function once on component mount
        initializeAndFetchData();

        return () => {
            isMounted = false; // Cleanup function to prevent state updates on unmounted component
        };
    }, []); // Empty dependency array ensures this runs once

    const mapWorkItemToProgress = (workItem: WorkItem): WorkItemProgress => {
        try {
            console.log(`Mapping work item ID: ${workItem.id}`);
            const fieldData = workItem.fields['Custom.AnswersField'];

            if (fieldData) {
                const decodedValue = decodeHtmlEntities(fieldData as string);
                const answers: { [key: string]: AnswerDetail } = JSON.parse(decodedValue);

                const totalEntries = answers ? Object.keys(answers).length : 0;
                const completedEntriesCount = totalEntries > 0 ?
                    Object.values(answers).filter(answer =>
                        answer.entries.every(entry => Boolean(entry.value))
                    ).length : 0;

                const progress = totalEntries > 0 ? (completedEntriesCount / totalEntries) * 100 : 0;

                return {
                    id: workItem.id,
                    title: workItem.fields['System.Title'] || "Untitled",
                    progress,
                    completed: completedEntriesCount,
                    total: totalEntries
                };
            } else {
                console.warn(`Work item ID ${workItem.id} does not have the 'Custom.AnswersField' field.`);
            }
        } catch (parseError) {
            console.error("Error parsing field data:", parseError);
        }

        return {
            id: workItem.id,
            title: workItem.fields['System.Title'] || "Untitled",
            progress: 0,
            completed: 0,
            total: 0
        };
    };

    if (loading) {
        return <div>Loading work item progress...</div>;
    }

    if (error) {
        return <div>Error: {error}</div>;
    }

    return (
        <div>
            <h3>Work Item Progress</h3>
            {workItemsProgress.map(({ id, title, progress, completed, total }) => (
                <div key={id}>
                    <h4>{title}</h4>
                    <div style={{ width: '100%', backgroundColor: '#e0e0e0', borderRadius: '8px', marginBottom: '10px' }}>
                        <div
                            style={{
                                width: `${progress}%`,
                                backgroundColor: '#76c7c0',
                                height: '24px',
                                borderRadius: '8px'
                            }}
                        />
                    </div>
                    <div>{`Work Item ID: ${id}, Progress: ${progress.toFixed(2)}%, Completed ${completed} out of ${total}`}</div>
                </div>
            ))}
        </div>
    );
};

showRootComponent(<QueryAndFieldDashboardWidget />, 'query-and-field-dashboard-root');
