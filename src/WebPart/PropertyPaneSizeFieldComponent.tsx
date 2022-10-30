import * as React from 'react';
import { Dropdown, IDropdownOption, Stack, TextField, Text } from '@fluentui/react';

export function PropertyPaneSizeFieldComponent(props: {
  value: string;
  setValue: (value: string) => void;
  label: string;
  placeholder: string;
  description: string;
  screenUnits: string;
}) {

  const screen = `v${props.screenUnits}`;

  const unitsOptions: IDropdownOption[] = [
    { key: screen, text: "% of the screen" },
    { key: '%', text: "% of the frame" },
    { key: 'cm', text: "centimeters" },
    { key: 'in', text: "inches" },
    { key: 'mm', text: "millimeters" },
    { key: 'pt', text: "points" },
    { key: 'px', text: "pixels" },
  ];

  const [value, setValue] = React.useState(props.value);
  React.useEffect(() => {
    const timeout = setTimeout(() => props.setValue(value), 500);
    return () => clearTimeout(timeout);
  }, [value]);

  const matches = value.match(/(\d+)\s*(\w+|%)?/);

  const placeholderMatches = props.placeholder.match(/(\d+)\s*(\w+|%)?/);

  const placeholderNumber = placeholderMatches?.[1];
  const placeholderUnits = placeholderMatches[2];

  const number = matches?.[1] ?? '';
  const units = matches?.[2] ?? (number ? screen : placeholderUnits);

  const onNumberChanged = (_, val) => {
    setValue(val ? val + units : '');
  };

  const onUnitChanged = (_, val) => {
    setValue(number + val.key);
  };

  return (
    <Stack tokens={{ childrenGap: "s2" }}>
      <Stack horizontal tokens={{ childrenGap: "s2" }}>
        <Stack.Item grow>
          <TextField label={props.label} placeholder={placeholderNumber} value={number} onChange={onNumberChanged} />
        </Stack.Item>
        <Stack.Item align='end'>
          <Dropdown style={{ minWidth: "10em" }} options={unitsOptions} selectedKey={units} disabled={number === ''} onChange={onUnitChanged} />
        </Stack.Item>
      </Stack>
      <Text variant='small' >{props.description}</Text>
    </Stack>
  );
}
