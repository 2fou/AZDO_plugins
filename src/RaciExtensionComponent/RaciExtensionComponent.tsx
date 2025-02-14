import * as React from 'react';
import { useState, useEffect } from 'react';
import * as SDK from 'azure-devops-extension-sdk';
import { IWorkItemFormService, WorkItemTrackingServiceIds } from 'azure-devops-extension-api/WorkItemTracking';
import { Dropdown } from 'azure-devops-ui/Dropdown';
import { DropdownSelection } from 'azure-devops-ui/Utilities/DropdownSelection';
import { Checkbox } from 'azure-devops-ui/Checkbox';
import { Button } from 'azure-devops-ui/Button';
import { Question, RoleAssignment, showRootComponent, decodeHtmlEntities } from '../Common/Common';
import { CommonServiceIds, IExtensionDataService } from 'azure-devops-extension-api';

const RaciExtensionComponent: React.FC = () => {
    const [workItemFormService, setWorkItemFormService] = useState<IWorkItemFormService | null>(null);
    const [questions, setQuestions] = useState<Question[]>([]);
    const [roles, setRoles] = useState<{ id: string, name: string }[]>([]);
    const [roleAssignments, setRoleAssignments] = useState<{
        [questionId: string]: {
            [entryLabel: string]: RoleAssignment[];
        };
    }>({});

useEffect(() => {
    const initializeComponent = async () => {
        await SDK.init();
        const wifService = await SDK.getService<IWorkItemFormService>(WorkItemTrackingServiceIds.WorkItemFormService);
        setWorkItemFormService(wifService);

        const extDataService = await SDK.getService<IExtensionDataService>(CommonServiceIds.ExtensionDataService);
        const dataManager = await extDataService.getExtensionDataManager(SDK.getExtensionContext().id, await SDK.getAccessToken());

        const rolesData = await dataManager.getValue<{ id: string, name: string }[]>('roles', { scopeType: 'Default' }) || [];
        setRoles(rolesData);
        console.log('Loaded roles:', rolesData);

        const questionVersions: Question[][] = await dataManager.getValue('questionVersions', { scopeType: 'Default' }) || [];
        const fieldValue = await wifService.getFieldValue('Custom.AnswersField', { returnOriginalValue: true }) as string;
        
        let parsedFieldData: { versionIndex: number, data: any } = { versionIndex: 0, data: {} };
        if (fieldValue) {
            const decodedFieldValue = decodeHtmlEntities(fieldValue);
            try {
                parsedFieldData = JSON.parse(decodedFieldValue);
            } catch (error) {
                console.error("Error parsing field value", error);
            }
        }

        const loadedQuestions = questionVersions[parsedFieldData.versionIndex] || [];
        setQuestions(loadedQuestions.filter(q => parsedFieldData.data[q.id]?.checked));

        try {
            const roleAssignmentsField = await wifService.getFieldValue('Custom.RoleAssignmentsField', { returnOriginalValue: true }) as string;
            if (roleAssignmentsField) {
                const decodedRoleAssignments = decodeHtmlEntities(roleAssignmentsField);
                const parsedRoleAssignments = JSON.parse(decodedRoleAssignments);
                setRoleAssignments(parsedRoleAssignments);
            }
        } catch (error) {
            console.error("Error loading role assignments from field:", error);
        }
    };

    initializeComponent();
}, []);

    const handleRoleAssignmentChange = (questionId: string, entryLabel: string, roleId: string, raci: string, index: number) => {
        setRoleAssignments(prev => {
            const newAssignments = [...(prev[questionId]?.[entryLabel] || [])];
            newAssignments[index] = { roleId, raci };
            const updatedAssignments = { ...prev, [questionId]: { ...prev[questionId], [entryLabel]: newAssignments } };
            saveRoleAssignments(updatedAssignments);
            return updatedAssignments;
        });
    };

    const addRoleAssignment = (questionId: string, entryLabel: string) => {
        setRoleAssignments(prev => {
            const updatedAssignments = {
                ...prev,
                [questionId]: {
                    ...prev[questionId],
                    [entryLabel]: [...(prev[questionId]?.[entryLabel] || []), { roleId: '', raci: '' }]
                }
            };
            saveRoleAssignments(updatedAssignments);
            return updatedAssignments;
        });
    };

    const removeRoleAssignment = (questionId: string, entryLabel: string, index: number) => {
        setRoleAssignments(prev => {
            const newAssignments = [...(prev[questionId]?.[entryLabel] || [])];
            newAssignments.splice(index, 1);
            const updatedAssignments = { ...prev, [questionId]: { ...prev[questionId], [entryLabel]: newAssignments } };
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
            <h2>RACI Role Assignment</h2>
            {questions.map(question => (
                <div key={question.id}>
                    <h3>{question.text}</h3>
                    {question.expectedEntries.labels.map((label, entryIndex) => (
                        <div key={entryIndex} style={{ marginBottom: '10px' }}>
                            <div>
                                <strong>{label}</strong>
                                {(roleAssignments[question.id]?.[label] || []).map((assignment, index) => {
                                    const role = roles.find(r => r.id === assignment.roleId);
                                    const dropdownSelection = new DropdownSelection();
                                    if (role) {
                                        dropdownSelection.select(roles.indexOf(role));
                                    }
                                    return (
                                        <div key={index} style={{ marginLeft: '20px' }}>
                                            <Dropdown
                                                items={roles.map(role => ({ id: role.id, text: role.name }))}
                                                selection={dropdownSelection}
                                                onSelect={(e, item) => handleRoleAssignmentChange(question.id, label, item.id, assignment.raci, index)}
                                                placeholder="Select Role"
                                            />
                                            {role && <span>Selected Role: {role.name}</span>}
                                            <Checkbox
                                                label="R"
                                                checked={assignment.raci.includes('R')}
                                                onChange={(e, checked) => {
                                                    const newRaci = checked ? `${assignment.raci}R` : assignment.raci.replace('R', '');
                                                    handleRoleAssignmentChange(question.id, label, assignment.roleId, newRaci, index);
                                                }}
                                            />
                                            <Checkbox
                                                label="A"
                                                checked={assignment.raci.includes('A')}
                                                onChange={(e, checked) => {
                                                    const newRaci = checked ? `${assignment.raci}A` : assignment.raci.replace('A', '');
                                                    handleRoleAssignmentChange(question.id, label, assignment.roleId, newRaci, index);
                                                }}
                                            />
                                            <Checkbox
                                                label="C"
                                                checked={assignment.raci.includes('C')}
                                                onChange={(e, checked) => {
                                                    const newRaci = checked ? `${assignment.raci}C` : assignment.raci.replace('C', '');
                                                    handleRoleAssignmentChange(question.id, label, assignment.roleId, newRaci, index);
                                                }}
                                            />
                                            <Checkbox
                                                label="I"
                                                checked={assignment.raci.includes('I')}
                                                onChange={(e, checked) => {
                                                    const newRaci = checked ? `${assignment.raci}I` : assignment.raci.replace('I', '');
                                                    handleRoleAssignmentChange(question.id, label, assignment.roleId, newRaci, index);
                                                }}
                                            />
                                            <Button text="Remove" onClick={() => removeRoleAssignment(question.id, label, index)} />
                                        </div>
                                    );
                                })}
                            </div>
                            <Button iconProps={{ iconName: 'Add' }} onClick={() => addRoleAssignment(question.id, label)}>Add Role</Button>
                        </div>
                    ))}
                </div>
            ))}
        </div>
    );
};

showRootComponent(<RaciExtensionComponent />, 'raci-extension-root');

export default RaciExtensionComponent;