import * as React from "react";
import { ISearchQuery, ISearchResult } from "@pnp/sp/search/types";
import { Dropdown, getFocusStyle, getTheme, IDropdownOption, ITheme, Link, List, mergeStyleSets, SearchBox, Shimmer, ShimmerElementsGroup, ShimmerElementType, Stack, TextField } from "@fluentui/react";
import { useSearch } from "pnp-react-hooks";

const disableWhen = (searchOptions: string | ISearchQuery) =>
{
    if (typeof searchOptions === "string")
    {
        return searchOptions.length < 1;
    }
    else if (searchOptions?.Querytext)
    {
        return searchOptions.Querytext.length < 1;
    }
    else
    {
        return true;
    }
};

const searchPageSize = 10;

export function BasicSearch()
{
    const [searchText, setSearchText] = React.useState<string>();

    const [searchResponse, setPage] = useSearch(
        {
            Querytext: searchText,
            RowsPerPage: searchPageSize
        },
        {
            disabled: disableWhen
        });

    return (
        <Stack tokens={{ childrenGap: 20 }}>
            <Stack.Item>
                <SearchBox
                    placeholder="Search on SharePoint"
                    onSearch={setSearchText}
                    onClear={() => setSearchText(undefined)}
                />
            </Stack.Item>
            <Stack.Item>
                {
                    searchText?.length > 0
                        ? <Shimmer
                            customElementsGroup={getCustomElements}
                            isDataLoaded={searchResponse !== undefined}
                        >
                            <Stack tokens={{ childrenGap: 15 }}>
                                <Stack.Item>
                                    <List
                                        items={searchResponse?.PrimarySearchResults ?? []}
                                        onRenderCell={onRenderCell}
                                    />
                                </Stack.Item>
                                <Stack.Item align="end">
                                    <Dropdown
                                        options={getPageOptions(searchResponse?.TotalRows, searchPageSize)}
                                        selectedKey={searchResponse?.CurrentPage}
                                        placeholder="Set page"
                                        styles={dropdownStyles}
                                        onChange={(_, opt) =>
                                        {
                                            if (typeof opt.key === "number") setPage(opt.key);
                                        }}
                                    />
                                </Stack.Item>
                            </Stack>
                        </Shimmer>
                        : undefined
                }
            </Stack.Item>
        </Stack>
    );
}

const getPageOptions = (totalRows: number, pageSize: number) =>
{
    const pageCount = Math.ceil(totalRows / pageSize);

    const options: IDropdownOption[] = [];

    for (let index = 1; index <= pageCount; index++)
        options.push({
            key: index,
            text: `Page ${index}`
        });

    return options;
};

const theme: ITheme = getTheme();
const { palette, semanticColors, fonts } = theme;

const dropdownStyles = { dropdown: { width: 100 } };
const classNames = mergeStyleSets({
    itemCell: [
        getFocusStyle(theme, { inset: -1 }),
        {
            minHeight: 54,
            padding: 10,
            boxSizing: 'border-box',
            borderBottom: `1px solid ${semanticColors.bodyDivider}`,
            display: 'flex',
            selectors: {
                '&:hover': { background: palette.neutralPrimaryAlt },
            },
        },
    ],
    itemContent: {
        marginLeft: 10,
        overflow: 'hidden',
        flexGrow: 1,
    },
    itemName: [
        fonts.xLarge,
        {
            whiteSpace: 'nowrap',
            overflow: 'hidden',
            textOverflow: 'ellipsis',
        },
    ],
    itemIndex: {
        fontSize: fonts.small.fontSize,
        color: palette.neutralTertiary,
        marginBottom: 10,
    }
});


const onRenderCell = (item: ISearchResult, index: number | undefined): JSX.Element =>
{
    return (
        <div className={classNames.itemCell} data-is-focusable={true}>
            <div className={classNames.itemContent}>
                <div className={classNames.itemName}>
                    <Link href={item.Path} target="_blank">
                        {item.Title}
                    </Link>
                </div>
                <div className={classNames.itemIndex}>{`View ${item.ViewsLifeTime}`}</div>
                <div>{item.Description}</div>
            </div>
        </div>
    );
};

const wrapperStyles = { display: 'flex' };


const getCustomElements = (): JSX.Element =>
{
    return (
        <div style={wrapperStyles}>
            <ShimmerElementsGroup
                flexWrap
                width="100%"
                shimmerElements={[
                    { type: ShimmerElementType.line, width: '100%', height: 50, verticalAlign: 'bottom' },
                    { type: ShimmerElementType.gap, width: '100%', height: 10, verticalAlign: 'bottom' },
                    { type: ShimmerElementType.line, width: '100%', height: 50, verticalAlign: 'bottom' },
                    { type: ShimmerElementType.gap, width: '100%', height: 10, verticalAlign: 'bottom' },
                    { type: ShimmerElementType.line, width: '100%', height: 50, verticalAlign: 'bottom' },
                    { type: ShimmerElementType.gap, width: '100%', height: 10, verticalAlign: 'bottom' },
                    { type: ShimmerElementType.line, width: '100%', height: 50, verticalAlign: 'bottom' },
                    { type: ShimmerElementType.gap, width: '100%', height: 10, verticalAlign: 'bottom' },
                    { type: ShimmerElementType.line, width: '100%', height: 50, verticalAlign: 'bottom' },
                ]}
            />
        </div>
    );
};