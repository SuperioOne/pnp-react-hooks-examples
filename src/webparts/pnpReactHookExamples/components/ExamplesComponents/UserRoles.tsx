import * as React from "react";
import { DetailsListLayoutMode, Dropdown, IColumn, IDropdownOption, ShimmeredDetailsList, Stack, Text } from "@fluentui/react";
import { IListInfo } from "@pnp/sp/lists/types";
import { IRoleAssignmentInfo, IRoleDefinitionInfo } from "@pnp/sp/security/types";
import { ISiteGroupInfo } from "@pnp/sp/site-groups";
import { ISiteUserInfo } from "@pnp/sp/site-users/types";
import { useRoleAssignments, useCurrentUser, useGroups, useLists } from "pnp-react-hooks";

// expanded version
export interface IRoleInfo extends IRoleAssignmentInfo
{
    Member: ISiteUserInfo | ISiteGroupInfo;
    RoleDefinitionBindings: IRoleDefinitionInfo[];
}

export interface IAssignedRole
{
    Member: ISiteUserInfo | ISiteGroupInfo;
    EffectiveRoleDefinition: IRoleDefinitionInfo;
}

export function UserRoles()
{
    const [selectedList, setSelectedList] = React.useState<IDropdownOption<IListInfo>>(null);

    const lists = useLists({
        query: {
            select: ["Title", "Id", "ItemCount"],
            orderBy: "Title",
            orderyByAscending: true
        }
    });

    const listOptions: IDropdownOption<IListInfo>[] = React.useMemo(() =>
    {
        return lists?.map(e => ({ key: e.Id, text: e.Title, data: e }));
    }, [lists]);

    const assignments = useRoleAssignments({
        query: {
            expand: ["RoleDefinitionBindings", "Member"],
            select: ["RoleDefinitionBindings", "PrincipalId"]
        },
        scope: selectedList
            ? { list: selectedList.data.Id }
            : undefined
    }) as IRoleInfo[];

    const user = useCurrentUser({
        query: {
            select: ["Id"]
        }
    });

    const userGroups = useGroups({
        query: {
            select: ["Id", "Title"]
        },
        userId: user?.Id,
        disabled: user === undefined || user.Id < 1
    });

    const userRoles: IAssignedRole[] = React.useMemo(() =>
    {
        if (assignments && user?.Id > 0 && userGroups)
        {
            const groupIds = new Set(userGroups.map(e => e.Id));

            return assignments
                .filter(e =>
                {
                    return e.PrincipalId === user.Id || groupIds.has(e.PrincipalId);
                })
                .map((e): IAssignedRole =>
                {
                    return {
                        Member: e.Member,
                        EffectiveRoleDefinition: e.RoleDefinitionBindings.sort(role => -role.Order)[0]
                    };
                });
        }
        else
        {
            return undefined;
        }
    }, [userGroups, assignments, user]);

    return (
        <Stack tokens={{ childrenGap: 15 }}>
            <Stack.Item>
                <Dropdown
                    options={listOptions}
                    disabled={!Array.isArray(lists)}
                    styles={_dropdownStyles}
                    label="Lists"
                    placeholder="Select a list"
                    selectedKey={selectedList?.key}
                    onChange={(_, opt) =>
                    {
                        setSelectedList(opt);
                    }}
                />
            </Stack.Item>
            <Stack.Item>
                <Text variant="xLargePlus" > {`Your roles on ${selectedList ? `'${selectedList.data.Title}'` : "current site"}.`} </Text>
            </Stack.Item>
            <Stack.Item>
                <ShimmeredDetailsList
                    items={userRoles ?? []}
                    columns={_colums}
                    layoutMode={DetailsListLayoutMode.fixedColumns}
                    enableShimmer={!Array.isArray(userRoles)}
                />
            </Stack.Item>
        </Stack >);
}

const _colums: IColumn[] =
    [
        {
            key: "from",
            name: "From",
            minWidth: 150,
            maxWidth: 150,
            isResizable: true,
            onRender: (item: IAssignedRole) => item.Member.Title
        },
        {
            key: "given",
            name: "Type",
            minWidth: 75,
            maxWidth: 75,
            isResizable: true,
            onRender: (item: IAssignedRole) => principalTypeMap[item.Member.PrincipalType] ?? "Unknow"
        },
        {
            key: "role",
            name: "Role",
            minWidth: 150,
            isResizable: true,
            onRender: (item: IAssignedRole) => item.EffectiveRoleDefinition.Name
        },
        {
            key: "description",
            name: "Description",
            minWidth: 250,
            isResizable: true,
            isMultiline: true,
            onRender: (item: IAssignedRole) => item.EffectiveRoleDefinition.Description
        }
    ];

const principalTypeMap = {
    0: "None",
    1: "User",
    2: "DistributionList",
    4: "SecurityGroup",
    8: "SharePointGroup",
    15: "All"
};

const _dropdownStyles = { dropdown: { width: 300 } };