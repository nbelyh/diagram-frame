import * as React from 'react';
import { Dropdown, IDropdownOption, Stack, TextField } from '@fluentui/react';

export function PropertyPaneSizeFieldComponent(props: {
  value: string;
  setValue: (value: string) => void;
  label: string;
  screenUnits: string;
}) {

  const unitsOptions: IDropdownOption[] = [
    { key: '', text: '' },
    { key: `v${props.screenUnits}`, text: '% of the screen' },
    { key: 'cm', text: 'centimeters' },
    { key: 'in', text: 'inches' },
    { key: 'mm', text: 'millimeters' },
    { key: 'pt', text: 'points' },
    { key: 'px', text: 'pixels' },
  ];

  const [value, setValue] = React.useState(props.value);
  React.useEffect(() => {
    const timeout = setTimeout(() => props.setValue(value), 500);
    return () => clearTimeout(timeout);
  }, [value]);

  const matches = value.match(/(\d+)\s*(\w+)?/);

  const number = matches?.[1] ?? '';
  const units = matches?.[2] ?? '';

  const onNumberChanged = (_, val) => {
    setValue(val ? val + units : '');
  };

  const onUnitChanged = (_, val) => {
    setValue(number + val.key);
  };

  return (
    <Stack horizontal tokens={{ childrenGap: "s2" }}>
      <Stack.Item grow>
        <TextField placeholder='auto' label={props.label} value={number} onChange={onNumberChanged} />
      </Stack.Item>
      <Stack.Item align='end'>
        <Dropdown placeholder='auto' style={{ minWidth: "10em" }} options={unitsOptions} selectedKey={units} onChange={onUnitChanged} />
      </Stack.Item>
    </Stack>
  );
}
