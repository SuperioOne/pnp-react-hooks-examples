import * as React from "react";
import * as dayjs from 'dayjs';
import { DefaultButton, DetailsListLayoutMode, Dropdown, IColumn, IDropdownOption, IStackItemStyles, MessageBar, MessageBarType, Separator, ShimmeredDetailsList, Stack, Text } from "@fluentui/react";
import { ListOptions, useList, useListItems, useLists } from "pnp-react-hooks";
import { IListInfo } from "@pnp/sp/lists/types";

export function PagedListItems()
{
    const [selectedList, setSelectedList] = React.useState<IDropdownOption<IListInfo>>(null);

    // get lists from current web
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

    return (
        <Stack tokens={{ childrenGap: 25 }}>
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
            <Stack.Item styles={_stackStyles}>
                <Separator />
            </Stack.Item>
            <Stack.Item>
                <SPPagedList
                    list={selectedList?.data.Id}
                />
            </Stack.Item>
        </Stack>
    );
}

export interface ISPListProps
{
    list?: string;
}

export function SPPagedList(props: ISPListProps)
{
    const [error, setError] = React.useState<Error>();
    const [loading, setLoading] = React.useState<boolean>(false);


    const [items, getNext, hasNext] = useListItems(props.list, {
        query: {
            select: ["Title", "Id", "Author/Title", "Editor/Title", "Modified", "Created"],
            expand: ["Author", "Editor"],
            top: 10
        },
        mode: ListOptions.Paged,
        error: setError
    });

    // get list title and Id to display header
    const list = useList(props.list, {
        query: {
            select: ["Title", "Id"]
        }
    });

    const onNext = React.useCallback(() =>
    {
        // disable "next" button
        setLoading(true);

        // get next with callback parameter to enable button back.
        getNext(() => setLoading(false));
    }, [getNext]);

    return (
        <Stack tokens={{ childrenGap: 5 }}>
            <Stack.Item>
                {
                    error ?
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
                            disabled={!hasNext || loading}
                            onClick={onNext}
                        >
                            Next
                        </DefaultButton>
                    </Stack.Item>
                </Stack>
            </Stack.Item>
            <Stack.Item>
                <ShimmeredDetailsList
                    items={items ?? []}
                    columns={_colums}
                    layoutMode={DetailsListLayoutMode.justified}
                    enableShimmer={props.list && !Array.isArray(items)}
                />
            </Stack.Item>
        </Stack>);
}

const _stackStyles: Partial<IStackItemStyles> = { root: { maxWidth: 300 } };
const _colums: IColumn[] =
    [
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

const _dropdownStyles = { dropdown: { width: 300 } };