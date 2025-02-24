import React, { useState, useEffect } from 'react';
import * as SDK from 'azure-devops-extension-sdk';
import { IWorkItemFormService, WorkItemTrackingServiceIds } from 'azure-devops-extension-api/WorkItemTracking';
import { Dropdown } from 'azure-devops-ui/Dropdown';
import { DropdownSelection } from 'azure-devops-ui/Utilities/DropdownSelection';
import { Checkbox } from 'azure-devops-ui/Checkbox';
import { Button } from 'azure-devops-ui/Button';
import { Deliverable, RoleAssignment, showRootComponent, decodeHtmlEntities, Question, Version, Role } from '../Common/Common';
import { CommonServiceIds, IExtensionDataService } from 'azure-devops-extension-api';



const RaciExtensionComponent: React.FC = () => {
    const [workItemFormService, setWorkItemFormService] = useState<IWorkItemFormService | null>(null);
    const [deliverables, setDeliverables] = useState<Deliverable[]>([]);
    const [selectedDeliverableIds, setSelectedDeliverableIds] = useState<Set<string>>(new Set());
    const [roles, setRoles] = useState<Role[]>([]);
    const [roleAssignments, setRoleAssignments] = useState<{ [deliverableId: string]: RoleAssignment[] }>({});
    const [currentVersionDescription, setCurrentVersionDescription] = useState<string>('');

    useEffect(() => {
        const initializeComponent = async () => {
            try {
                console.log("Initializing component...");
                await SDK.init();
                const wifService = await SDK.getService<IWorkItemFormService>(WorkItemTrackingServiceIds.WorkItemFormService);
                setWorkItemFormService(wifService);
    
                const extDataService = await SDK.getService<IExtensionDataService>(CommonServiceIds.ExtensionDataService);
                const dataManager = await extDataService.getExtensionDataManager(SDK.getExtensionContext().id, await SDK.getAccessToken());
    
                // Fetch roles and deliverables
                const rolesData = await dataManager.getValue<Role[]>('roles', { scopeType: 'Default' }) || [];
                console.log("Fetched roles:", rolesData);
                setRoles(rolesData);
    
                const storedDeliverables = await dataManager.getValue<Deliverable[]>('deliverables', { scopeType: 'Default' }) || [];
                console.log("Fetched deliverables:", storedDeliverables);
                setDeliverables(storedDeliverables);
    
                // Fetch versions of questions
                const versions: Version[] = await dataManager.getValue<Version[]>('questionaryVersions', { scopeType: 'Default' }) || [];
                console.log("Fetched questionary versions:", versions);
    
                // Fetch and process answers field
                const answersFieldValue = await wifService.getFieldValue('Custom.AnswersField', { returnOriginalValue: true }) as string;
                console.log("Fetched answers field value:", answersFieldValue);
    
                if (answersFieldValue) {
                    const decodedAnswersField = decodeHtmlEntities(answersFieldValue);
                    const parsedAnswersField = JSON.parse(decodedAnswersField);
                    const targetVersion = parsedAnswersField.version || '';
                    console.log("Parsed answers field:", parsedAnswersField);
    
                    // Match the version description
                    const matchedVersion = versions.find(v => v.description === targetVersion);
                    if (matchedVersion) {
                        const selectedQuestionIds: string[] = parsedAnswersField.selectedQuestions || [];
                        console.log("Using questions from matched version:", matchedVersion.description);
    
                        if (selectedQuestionIds.length > 0) {
                            const deliverableIds = new Set<string>();
                            matchedVersion.questions.forEach(({ id, linkedDeliverables }) => {
                                if (selectedQuestionIds.includes(id)) {
                                    linkedDeliverables.forEach(deliverableId => {
                                        deliverableIds.add(deliverableId);
                                    });
                                }
                            });
                            console.log("Selected deliverable IDs:", deliverableIds);
                            setSelectedDeliverableIds(deliverableIds);
                        }
                    }
                }
    
                // Load role assignments
                const roleAssignmentsField = await wifService.getFieldValue('Custom.RoleAssignmentsField', { returnOriginalValue: true }) as string;
                console.log("Fetched role assignments field:", roleAssignmentsField);
    
                if (roleAssignmentsField) {
                    const decodedRoleAssignments = decodeHtmlEntities(roleAssignmentsField);
                    const parsedRoleAssignments = JSON.parse(decodedRoleAssignments);
                    console.log("Parsed role assignments:", parsedRoleAssignments);
                    setRoleAssignments(parsedRoleAssignments);
                }
            } catch (error) {
                console.error("Initialization error:", error);
            }
        };
    
        initializeComponent();
    }, []);

    const handleRoleAssignmentChange = (deliverableId: string, roleId: string, raci: string, index: number) => {
        setRoleAssignments(prev => {
            const newAssignments = [...(prev[deliverableId] || [])];
            newAssignments[index] = { roleId, raci };
            const updatedAssignments = { ...prev, [deliverableId]: newAssignments };
            saveRoleAssignments(updatedAssignments);
            return updatedAssignments;
        });
    };

    const addRoleAssignment = (deliverableId: string) => {
        setRoleAssignments(prev => {
            const updatedAssignments = {
                ...prev,
                [deliverableId]: [...(prev[deliverableId] || []), { roleId: '', raci: '' }]
            };
            saveRoleAssignments(updatedAssignments);
            return updatedAssignments;
        });
    };

    const removeRoleAssignment = (deliverableId: string, index: number) => {
        setRoleAssignments(prev => {
            const newAssignments = [...(prev[deliverableId] || [])];
            newAssignments.splice(index, 1);
            const updatedAssignments = { ...prev, [deliverableId]: newAssignments };
            saveRoleAssignments(updatedAssignments);
            return updatedAssignments;
        });
    };

    const saveRoleAssignments = async (assignments: any) => {
        try {
            if (workItemFormService) {
                const serializedData = JSON.stringify(assignments);
                await workItemFormService.setFieldValue('Custom.RoleAssignmentsField', serializedData);
            }
        } catch (error) {
            console.error("Error saving role assignments:", error);
            alert("Failed to save roles and RACI assignments.");
        }
    };

    return (
        <div>
            <h2>RACI Role Assignment for Deliverables</h2>
            <p>Version: {currentVersionDescription}</p>
            {deliverables.filter(deliverable => selectedDeliverableIds.has(deliverable.id)).map(deliverable => (
                <div key={deliverable.id}>
                    <h3>{deliverable.label}</h3>
                    {(roleAssignments[deliverable.id] || []).map((assignment, index) => {
                        const role = roles.find(r => r.id === assignment.roleId);
                        const dropdownSelection = new DropdownSelection();
                        if (role) {
                            dropdownSelection.select(roles.indexOf(role));
                        }
                        return (
                            <div key={index} style={{ marginLeft: '20px', marginBottom: '10px' }}>
                                <Dropdown
                                    items={roles.map(role => ({ id: role.id, text: role.name }))}
                                    selection={dropdownSelection}
                                    onSelect={(e, item) => handleRoleAssignmentChange(deliverable.id, item.id, assignment.raci, index)}
                                    placeholder="Select Role"
                                />
                                {['R', 'A', 'C', 'I'].map(raciChar => (
                                    <Checkbox
                                        key={raciChar}
                                        label={raciChar}
                                        checked={assignment.raci.includes(raciChar)}
                                        onChange={(e, checked) => {
                                            const newRaci = checked ? `${assignment.raci}${raciChar}` : assignment.raci.replace(raciChar, '');
                                            handleRoleAssignmentChange(deliverable.id, assignment.roleId, newRaci, index);
                                        }}
                                    />
                                ))}
                                <Button text="Remove" onClick={() => removeRoleAssignment(deliverable.id, index)} />
                            </div>
                        );
                    })}
                    <Button iconProps={{ iconName: 'Add' }} onClick={() => addRoleAssignment(deliverable.id)}>Add Role</Button>
                </div>
            ))}
        </div>
    );
};

showRootComponent(<RaciExtensionComponent />, 'raci-extension-root');

export default RaciExtensionComponent;