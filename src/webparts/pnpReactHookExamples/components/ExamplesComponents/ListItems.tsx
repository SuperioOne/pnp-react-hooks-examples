import * as React from "react";
import * as dayjs from 'dayjs';
import { DetailsListLayoutMode, Dropdown, IColumn, IDropdownOption, IStackItemStyles, MessageBar, MessageBarType, Separator, ShimmeredDetailsList, Slider, Stack, Text } from "@fluentui/react";
import { useList, useListItems, useLists } from "pnp-react-hooks";
import { IListInfo } from "@pnp/sp/lists/types";

export function ListItems()
{
    const rangeRef = React.useRef<number>();
    const [sliderValue, setSliderValue] = React.useState(10);
    const [sliderLowerValue, setSliderLowerValue] = React.useState(0);
    const [range, setRange] = React.useState<[number, number]>([sliderLowerValue, sliderValue]);
    const [selectedList, setSelectedList] = React.useState<IDropdownOption<IListInfo>>(null);

    // get lists from current web
    const lists = useLists({
        query: {
            select: ["Title", "Id", "ItemCount"],
            orderBy: "Title",
            orderyByAscending: true
        }
    });

    const onChange = (_: unknown, newRange: [number, number]) =>
    {
        if (rangeRef.current > 0)
            clearTimeout(rangeRef.current);

        setSliderLowerValue(newRange[0]);
        setSliderValue(newRange[1]);

        rangeRef.current = setTimeout(() =>
        {
            setRange(newRange);
            
            if (rangeRef.current > 0)
                clearTimeout(rangeRef.current);
        }, 1000);
    };

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
                <Slider
                    label="Item Range"
                    lowerValue={sliderLowerValue}
                    max={selectedList?.data?.ItemCount ?? 100}
                    min={0}
                    step={1}
                    onChange={onChange}
                    ranged
                    showValue
                    value={sliderValue}
                />
            </Stack.Item>
            <Stack.Item styles={_stackStyles}>
                <Separator />
            </Stack.Item>
            <Stack.Item>
                <SPList
                    list={selectedList?.data.Id}
                    skip={range[0]}
                    top={range[1] - range[0]}
                />
            </Stack.Item>
        </Stack>
    );
}

export interface ISPListProps
{
    list?: string;
    skip?: number;
    top?: number;
}

export function SPList(props: ISPListProps)
{
    const [error, setError] = React.useState<Error>();

    const items = useListItems(props.list, {
        query: {
            select: ["Title", "Id", "Author/Title", "Editor/Title", "Modified", "Created"],
            expand: ["Author", "Editor"],
            skip: props.skip,
            top: props.top,
        },
        exception: setError
    });

    const list = useList(props.list, {
        query: {
            select: ["Title", "Id"]
        }
    });

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