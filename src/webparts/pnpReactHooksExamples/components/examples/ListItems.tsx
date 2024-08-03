/* eslint-disable @typescript-eslint/no-non-null-assertion */
import * as React from "react";
import * as dayjs from 'dayjs';
import { DetailsListLayoutMode, Dropdown, IColumn, IDropdownOption, IStackItemStyles, MessageBar, MessageBarType, Separator, ShimmeredDetailsList, Slider, Stack, Text } from "@fluentui/react";
import { useList, useListItems, useLists } from "pnp-react-hooks";
import { IListInfo } from "@pnp/sp/lists/types";

const STACK_STYLES: Partial<IStackItemStyles> = { root: { maxWidth: 300 } };
const DROPDOWN_STYLES = { dropdown: { width: 300 } };
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

export interface ISPListProps {
    list?: string;
    skip?: number;
    top?: number;
}

export default function ListItems(): JSX.Element {
    const rangeRef = React.useRef<number>(0);
    const [sliderValue, setSliderValue] = React.useState(10);
    const [sliderLowerValue, setSliderLowerValue] = React.useState(0);
    const [range, setRange] = React.useState<[number, number]>([sliderLowerValue, sliderValue]);
    const [selectedList, setSelectedList] = React.useState<IDropdownOption<IListInfo> | undefined>(undefined);

    // get lists from current web
    const lists = useLists({
        query: {
            select: ["Title", "Id", "ItemCount"],
            orderBy: "Title",
            orderyByAscending: true
        }
    });

    const onChange = (_: unknown, newRange: [number, number]): void => {
        if ((rangeRef.current ?? 0) > 0)
            clearTimeout(rangeRef.current);

        setSliderLowerValue(newRange[0]);
        setSliderValue(newRange[1]);

        rangeRef.current = window.setTimeout(() => {
            setRange(newRange);

            if (rangeRef.current > 0)
                clearTimeout(rangeRef.current);
        }, 1000);
    };

    const listOptions: IDropdownOption<IListInfo>[] = React.useMemo(() => {
        return lists?.map(e => ({ key: e.Id, text: e.Title, data: e })) ?? [];
    }, [lists]);

    return (
        <Stack tokens={{ childrenGap: 25 }}>
            <Stack.Item>
                <Dropdown
                    options={listOptions}
                    disabled={!Array.isArray(lists)}
                    styles={DROPDOWN_STYLES}
                    label="Lists"
                    placeholder="Select a list"
                    selectedKey={selectedList?.key}
                    onChange={(_, opt) => {
                        setSelectedList(opt);
                    }}
                />
            </Stack.Item>
            <Stack.Item styles={STACK_STYLES}>
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
            <Stack.Item styles={STACK_STYLES}>
                <Separator />
            </Stack.Item>
            <Stack.Item>
                <SPList
                    list={selectedList?.data?.Id}
                    skip={range[0]}
                    top={range[1] - range[0]}
                />
            </Stack.Item>
        </Stack>
    );
}

export function SPList(props: ISPListProps): JSX.Element {
    const [error, setError] = React.useState<Error>();

    const items = useListItems(props.list!, {
        query: {
            select: ["Title", "Id", "Author/Title", "Editor/Title", "Modified", "Created"],
            expand: ["Author", "Editor"],
            skip: props.skip,
            top: props.top,
        },
        error: setError
    });

    const list = useList(props.list!, {
        query: {
            select: ["Title", "Id"]
        }
    });

    if (props.list) {
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
                        columns={COLUMNS}
                        layoutMode={DetailsListLayoutMode.justified}
                        enableShimmer={!items}
                    />
                </Stack.Item>
            </Stack>);
    }
    else {

        return (<div />);
    }
}


