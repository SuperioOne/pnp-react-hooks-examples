/* eslint-disable @typescript-eslint/no-non-null-asserted-optional-chain */
/* eslint-disable @typescript-eslint/no-non-null-assertion */
import * as React from "react";
import { DetailsListLayoutMode, getTheme, IColumn, INavLink, INavStyles, Nav, ShimmeredDetailsList, Stack, Text } from "@fluentui/react";
import { ISiteGroupInfo } from "@pnp/sp/site-groups/types";
import { useGroups, useGroupUsers, useIsMemberOf } from "pnp-react-hooks";

const THEME = getTheme();
const NAV_STYLES: Partial<INavStyles> = {
    root: {
        boxSizing: 'border-box',
        borderRight: '1px solid #333333',
        overflowY: 'auto',
        paddingRight: 5
    },
};

const COLUMNS: IColumn[] = [
    {
        key: "Id",
        name: "Id",
        minWidth: 20,
        maxWidth: 50,
        isResizable: true,
        fieldName: "Id"
    },
    {
        key: "Title",
        name: "Title",
        minWidth: 70,
        isResizable: true,
        fieldName: "Title"
    },
    {
        key: "Email",
        name: "Email",
        minWidth: 100,
        isResizable: true,
        fieldName: "Email"
    }
];

export default function GroupsAndUsers(): JSX.Element {
    const [selectedGroup, setGroup] = React.useState<ISiteGroupInfo>();
    const groups = useGroups({
        query: {
            select: ["*", "Title", "Id", "Description"],
            orderBy: "Title",
            orderyByAscending: true,
            filter: "IsHiddenInUI eq false and OwnerTitle ne 'System Account'"
        }
    });

    const groupLinks: INavLink[] = React.useMemo(() => {
        return groups?.map(e => ({
            name: e.Title,
            key: e.Id.toString(),
            onClick: () => setGroup(e),
            url: ""
        })) ?? [];
    }, [groups]);

    const users = useGroupUsers(selectedGroup?.Id!, {
        query: {
            select: ["*", "Id", "Email", "LoginName", "Title"],
            top: 100,
            orderBy: "Title",
            orderyByAscending: true
        }
    });

    const [amIMemberOf] = useIsMemberOf(selectedGroup?.Id!);

    return (
        <Stack horizontal verticalAlign="start" tokens={{ childrenGap: 15 }}>
            <Stack.Item>
                <Nav
                    selectedKey={selectedGroup?.Id.toString()}
                    styles={NAV_STYLES}
                    groups={[{
                        links: groupLinks,
                        name: "Sharpeoint Groups",
                    }]}
                />
            </Stack.Item>
            <Stack.Item style={{ width: "100%" }}>
                <Stack tokens={{ childrenGap: 15 }}>
                    <Stack.Item>
                        <Text variant="xLargePlus">{selectedGroup?.Title}</Text>
                    </Stack.Item>
                    <Stack.Item>
                        <Text variant="small">{selectedGroup?.Description}</Text>
                    </Stack.Item>
                    <Stack.Item>
                        <Text variant="mediumPlus" style={{ color: THEME.semanticColors.warningIcon }}>
                            {
                                amIMemberOf === true ? `You are a member of '${selectedGroup?.Title}'.`
                                    : amIMemberOf === false ? `You are not a member of '${selectedGroup?.Title}'.`
                                        : ""
                            }
                        </Text>
                    </Stack.Item>
                    <Stack.Item>
                        {selectedGroup
                            ? <ShimmeredDetailsList
                                items={users ?? []}
                                columns={COLUMNS}
                                layoutMode={DetailsListLayoutMode.justified}
                                enableShimmer={!Array.isArray(users)}
                            />
                            : undefined
                        }
                    </Stack.Item>
                </Stack>
            </Stack.Item>
        </Stack>);
}


