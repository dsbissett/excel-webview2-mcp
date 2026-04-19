export enum ToolCategory {
  INPUT = 'input',
  NAVIGATION = 'navigation',
  EXCEL = 'excel',
  EMULATION = 'emulation',
  PERFORMANCE = 'performance',
  NETWORK = 'network',
  DEBUGGING = 'debugging',
  IN_PAGE = 'in-page',
  LIFECYCLE = 'lifecycle',
}

export const labels = {
  [ToolCategory.INPUT]: 'Input automation',
  [ToolCategory.NAVIGATION]: 'Navigation automation',
  [ToolCategory.EXCEL]: 'Excel',
  [ToolCategory.EMULATION]: 'Emulation',
  [ToolCategory.PERFORMANCE]: 'Performance',
  [ToolCategory.NETWORK]: 'Network',
  [ToolCategory.DEBUGGING]: 'Debugging',
  [ToolCategory.IN_PAGE]: 'In-page tools',
  [ToolCategory.LIFECYCLE]: 'Add-in lifecycle',
};
