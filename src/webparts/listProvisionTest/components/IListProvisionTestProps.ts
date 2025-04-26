export interface IListProvisionTestProps {
  projectsListExists: boolean;
  onConfigureClick: () => void;
  userDisplayName: string;
  items?: {
    ID: string;
    Title: string;
    Status: string;
    AssignedTo: {
      Title: string;
    };
  }[];
}


