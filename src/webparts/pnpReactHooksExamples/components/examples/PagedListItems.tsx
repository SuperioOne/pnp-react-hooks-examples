/* eslint-disable @typescript-eslint/no-non-null-assertion */
import * as React from "react";
import * as dayjs from "dayjs";
import { Dropdown, Text, IColumn, IDropdownOption, IStackItemStyles, MessageBar, MessageBarType, Separator, Stack, DefaultButton, ShimmeredDetailsList, DetailsListLayoutMode } from "@fluentui/react";
import { ListItemsMode, useList, useListItems, useLists } from "pnp-react-hooks";
import { IListInfo } from "@pnp/sp/lists/types";

const DROPDOWN_STYLES = { dropdown: { width: 300 } };
const STACK_STYLES: Partial<IStackItemStyles> = { root: { maxWidth: 300 } };
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
        minWidth: 100,
        isResizable: true,
        fieldName: "Title"
    },
    {
        key: "Created By",
        name: "Created By",
        minWidth: 100,
        isResizable: true,
        onRender: item => item?.Author?.Title
    },
    {
        key: "Modified By",
        name: "Modified By",
        minWidth: 100,
        isResizable: true,
        onRender: item => item?.Editor?.Title
    },
    {
        key: "Created",
        name: "Created",
        minWidth: 100,
        isResizable: true,
        onRender: item => dayjs(item?.Created).format('DD/MM/YYYY')
    },
    {
        key: "Modified",
        name: "Modified",
        minWidth: 100,
        isResizable: true,
        onRender: item => dayjs(item?.Modified).format('DD/MM/YYYY')
    }
];

export default function PagedListItems(): JSX.Element {
    const [selectedList, setSelectedList] = React.useState<IDropdownOption<IListInfo>>();

    // get lists from current web
    const lists = useLists({
        query: {
            select: ["Title", "Id", "ItemCount"],
            orderBy: "Title",
            orderyByAscending: true
        }
    });

    const listOptions: IDropdownOption<IListInfo>[] = React.useMemo(() => {
        return lists?.map(e => ({ key: e.Id, text: e.Title, data: e })) ?? [];
    }, [lists]);

    const listId = selectedList?.key.toString() ?? "";

    return (
        <Stack tokens={{ childrenGap: 25 }}>
            <Stack.Item>
                <Dropdown
                    options={listOptions}
                    disabled={!lists}
                    styles={DROPDOWN_STYLES}
                    label="Lists"
                    placeholder="Select a list"
                    selectedKey={selectedList?.key}
                    onChange={(_, opt) => setSelectedList(opt)}
                />
            </Stack.Item>
            <Stack.Item styles={STACK_STYLES}>
                <Separator />
            </Stack.Item>
            <Stack.Item>
                <SPPagedList list={listId} />
            </Stack.Item>
        </Stack>
    );
}

interface ISPListProps {
    list: string;
}

function SPPagedList(props: ISPListProps): JSX.Element {
    const [error, setError] = React.useState<Error>();
    const [loading, setLoading] = React.useState<boolean>(false);

    // get list as pages
    const [page, getNext, done] = useListItems(props.list, {
        query: {
            select: ["Title", "Id", "Author/Title", "Editor/Title", "Modified", "Created"],
            expand: ["Author", "Editor"],
            top: 10
        },
        mode: ListItemsMode.Paged,
        error: setError,
    });

    // get list title and Id for header
    const list = useList(props.list!, {
        query: {
            select: ["Title", "Id"]
        }
    });

    const onNext = React.useCallback(() => {
        // disable "next" button
        setLoading(true);
        // get next with callback parameter to enable button back.
        getNext(() => setLoading(false));
    }, [getNext]);

    if (props.list.length < 1) {
        return (<div> </div>);
    }
    else {
        return (
            <Stack tokens={{ childrenGap: 5 }}>
                <Stack.Item>
                    {error ?
                        <MessageBar
                            messageBarType={MessageBarType.error}
                            isMultiline={true}
                            onDismiss={() => setError(undefined)}
                        >
                            <Stack tokens={{ childrenGap: 3 }}>
                                <Stack.Item>
                                    {error.message}
                                </Stack.Item>
                                <Stack.Item>
                                    {error.stack ?? ""}
                                </Stack.Item>
                            </Stack>
                        </MessageBar>
                        : undefined
                    }
                </Stack.Item>
                <Stack.Item>
                    <Text variant="xLargePlus">{list?.Title ?? ""}</Text>
                </Stack.Item>
                <Stack.Item>
                    <Text variant="small">{list?.Id ?? ""}</Text>
                </Stack.Item>
                <Stack.Item>
                    <Stack horizontal>
                        <Stack.Item>
                            <DefaultButton
                                disabled={done || loading}
                                onClick={onNext}
                            >
                                Next
                            </DefaultButton>
                        </Stack.Item>
                    </Stack>
                </Stack.Item>
                <Stack.Item>
                    <ShimmeredDetailsList
                        items={page ?? []}
                        columns={COLUMNS}
                        layoutMode={DetailsListLayoutMode.justified}
                        enableShimmer={!page}
                    />
                </Stack.Item>
            </Stack>);
    }
}


