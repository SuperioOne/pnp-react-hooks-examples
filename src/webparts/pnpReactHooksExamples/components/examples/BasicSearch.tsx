import * as React from "react";
import { ISearchQuery, ISearchResult } from "@pnp/sp/search/types";
import { Dropdown, getFocusStyle, getTheme, IDropdownOption, ITheme, Link, List, mergeStyleSets, SearchBox, Shimmer, ShimmerElementsGroup, ShimmerElementType, Stack } from "@fluentui/react";
import { useSearch } from "pnp-react-hooks";

const PAGE_SIZE = 10;
const THEME: ITheme = getTheme();
const { palette, semanticColors, fonts } = THEME;
const DROPDOWN_STYLES = { dropdown: { width: 100 } };
const WRAPPER_STYLES = { display: 'flex' };
const CLASS_NAMES = mergeStyleSets({
  itemCell: [
    getFocusStyle(THEME, { inset: -1 }),
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

const getPageOptions = (totalRows: number, pageSize: number): IDropdownOption[] => {
  const pageCount = Math.ceil(totalRows / pageSize);
  const options: IDropdownOption[] = [];

  for (let index = 1; index <= pageCount; index++)
    options.push({
      key: index,
      text: `Page ${index}`
    });

  return options;
};

const onRenderCell = (item: ISearchResult): JSX.Element => {
  return (
    <div className={CLASS_NAMES.itemCell} data-is-focusable={true}>
      <div className={CLASS_NAMES.itemContent}>
        <div className={CLASS_NAMES.itemName}>
          <Link href={item.Path} target="_blank">
            {item.Title}
          </Link>
        </div>
        <div className={CLASS_NAMES.itemIndex}>{`View ${item.ViewsLifeTime}`}</div>
        <div>{item.Description}</div>
      </div>
    </div>
  );
};

const getCustomElements = (): JSX.Element => {
  return (
    <div style={WRAPPER_STYLES}>
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

// example custom disable check function for disabling search when input text is empty.
const disableWhen = (searchOptions: string | ISearchQuery): boolean => {
  if (typeof searchOptions === "string") {
    return searchOptions.length < 1;
  }
  else if (searchOptions?.Querytext) {
    return searchOptions.Querytext.length < 1;
  }
  else {
    return true;
  }
};


export default function BasicSearch(): JSX.Element {
  const [searchText, setSearchText] = React.useState<string>();
  const [searchResponse, setPage] = useSearch(
    {
      Querytext: searchText,
      RowsPerPage: PAGE_SIZE
    },
    {
      disabled: disableWhen
    },
    [searchText]);

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
          (searchText?.length ?? 0) > 0
            ? <Shimmer
              customElementsGroup={getCustomElements}
              isDataLoaded={searchResponse !== undefined}
            >
              <Stack tokens={{ childrenGap: 15 }}>
                <Stack.Item>
                  <List
                    items={searchResponse?.primarySearchResults ?? []}
                    onRenderCell={onRenderCell}
                  />
                </Stack.Item>
                <Stack.Item align="end">
                  <Dropdown
                    options={getPageOptions((searchResponse?.totalRows ?? 0), PAGE_SIZE)}
                    selectedKey={searchResponse?.currentPage}
                    placeholder="Set page"
                    styles={DROPDOWN_STYLES}
                    onChange={(_, opt) => {
                      if (opt && typeof opt.key === "number") {
                        setPage(opt.key);
                      }
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


