import * as React from "react";
import { getFocusStyle, getTheme, ITheme, Link, List, mergeStyleSets, SearchBox, Shimmer, ShimmerElementsGroup, ShimmerElementType, Stack } from "@fluentui/react";
import { useSearchUser } from "pnp-react-hooks";
import { IPeoplePickerEntity } from "@pnp/sp/profiles/types";

export function BasicUserSearch()
{
    const [searchText, setSearchText] = React.useState<string>();

    const users = useSearchUser(
        {
            AllowEmailAddresses: true,
            MaximumEntitySuggestions: 5,
            QueryString: searchText
        });

    return (
        <Stack tokens={{ childrenGap: 20 }}>
            <Stack.Item>
                <SearchBox
                    placeholder="Search users"
                    onSearch={setSearchText}
                    onClear={() => setSearchText(undefined)}
                />
            </Stack.Item>
            <Stack.Item>
                {
                    searchText?.length > 0
                        ? <Shimmer
                            customElementsGroup={getCustomElements}
                            isDataLoaded={users !== undefined}
                        >
                            <Stack tokens={{ childrenGap: 15 }}>
                                <Stack.Item>
                                    <List
                                        items={users ?? []}
                                        onRenderCell={onRenderCell}
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

const theme: ITheme = getTheme();
const { palette, semanticColors, fonts } = theme;

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


const onRenderCell = (item: IPeoplePickerEntity, index: number | undefined): JSX.Element =>
{
    return (
        <div className={classNames.itemCell} data-is-focusable={true}>
            <div className={classNames.itemContent}>
                <div className={classNames.itemName}>
                    <Link href={`mailto:${item.EntityData.Email}`} target="_blank">
                        {item.DisplayText}
                    </Link>
                </div>
                <div className={classNames.itemIndex}>{`View ${item.EntityData.Department}`}</div>
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