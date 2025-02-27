import React, { useEffect, useState } from 'react';
import * as SDK from 'azure-devops-extension-sdk';
import { IWorkItemFormService, WorkItemTrackingServiceIds } from 'azure-devops-extension-api/WorkItemTracking';
import { CommonServiceIds, IExtensionDataService } from 'azure-devops-extension-api';
import { Deliverable, Role, decodeHtmlEntities, showRootComponent } from '../Common/Common';
import './RaciTableComponent.scss'; 
import { Tooltip } from 'react-tooltip';

const RaciTableComponent: React.FC = () => {
  const [deliverables, setDeliverables] = useState<Deliverable[]>([]);
  const [roles, setRoles] = useState<Role[]>([]);
  const [roleAssignments, setRoleAssignments] = useState<{ [deliverableId: string]: { roleId: string, raci: string }[] }>({});

  useEffect(() => {
    const initializeComponent = async () => {
      try {
        await SDK.init();
        const wifService = await SDK.getService<IWorkItemFormService>(WorkItemTrackingServiceIds.WorkItemFormService);
        const extDataService = await SDK.getService<IExtensionDataService>(CommonServiceIds.ExtensionDataService);
        const dataManager = await extDataService.getExtensionDataManager(SDK.getExtensionContext().id, await SDK.getAccessToken());

        const rolesData = await dataManager.getValue<Role[]>('roles', { scopeType: 'Default' }) || [];
        setRoles(rolesData);

        const storedDeliverables = await dataManager.getValue<Deliverable[]>('deliverables', { scopeType: 'Default' }) || [];
        setDeliverables(storedDeliverables);

        const roleAssignmentsField = await wifService.getFieldValue('Custom.RoleAssignmentsField', { returnOriginalValue: true }) as string;
        if (roleAssignmentsField) {
          const decodedAssignments = decodeHtmlEntities(roleAssignmentsField);
          setRoleAssignments(JSON.parse(decodedAssignments));
        }
      } catch (error) {
        console.error("Initialization error:", error);
      }
    };
    initializeComponent();
  }, []);

  const renderMatrix = () => {
    const deliverableIds = Object.keys(roleAssignments);
    const relevantDeliverables = deliverables.filter(deliverable => deliverableIds.includes(deliverable.id));
    const roleIdsSet = new Set<string>();

    Object.values(roleAssignments).forEach(assignments => {
      assignments.forEach(assignment => roleIdsSet.add(assignment.roleId));
    });

    const relevantRoles = roles.filter(role => roleIdsSet.has(role.id));

    return (
      <table className="raci-table">
        <thead>
          <tr>
            <th>Deliverable</th>
            {relevantRoles.map(role => (
              <th key={role.id} data-tooltip-id={`role-${role.id}`}>
                {role.name}
                <Tooltip id={`role-${role.id}`} place="top">
                  <div>
                    <strong>Description:</strong> {role.description}
                    <br />
                    <strong>Email:</strong> {role.email}
                    <br />
                    <strong>Department:</strong> {role.department}
                  </div>
                </Tooltip>
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {relevantDeliverables.map(deliverable => (
            <tr key={deliverable.id}>
              <td>{deliverable.label}</td>
              {relevantRoles.map(role => {
                const roleAssignment = roleAssignments[deliverable.id]?.find(ra => ra.roleId === role.id);
                return (
                  <td key={role.id}>
                    {roleAssignment
                      ? roleAssignment.raci.split('').map((duty, index) => (
                          <span key={index} className={`duty-${duty.toLowerCase()}`}>
                            {duty}
                          </span>
                        ))
                      : '-'}
                  </td>
                );
              })}
            </tr>
          ))}
        </tbody>
      </table>
    );
  };

  return (
    <div>
      <h2>RACI Matrix</h2>
      {renderMatrix()}
    </div>
  );
};

showRootComponent(<RaciTableComponent />, 'raci-table-root');
export default RaciTableComponent;