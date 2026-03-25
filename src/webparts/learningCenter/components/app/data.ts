export const companies = [
    { id: 'microsoft', name: 'Microsoft', logoUrl: 'https://upload.wikimedia.org/wikipedia/commons/4/44/Microsoft_logo.svg' },
    { id: 'google', name: 'Google', logoUrl: 'https://upload.wikimedia.org/wikipedia/commons/2/2f/Google_2015_logo.svg' },
    { id: 'aws', name: 'AWS', logoUrl: 'https://upload.wikimedia.org/wikipedia/commons/9/93/Amazon_Web_Services_Logo.svg' },
];

export const microsoftCertifications = [
    {
        category: 'Calling for Microsoft Teams',
        level: 'Teams Phone Specialization',
        url: 'https://learn.microsoft.com/en-us/credentials/specializations/microsoft-teams/teams-phone-specialization',
        certs: [
            {
                id: 'ms-700',
                name: 'Teams Administrator Associate',
                code: 'MS-700',
                url: 'https://learn.microsoft.com/en-us/credentials/certifications/m365-teams-administrator-associate/',
                isMandatory: true,
                minRequired: 1,
                duration: '12-14 hours',
                estimatedCompletion: '2 weeks',
                prerequisites: 'Foundational Knowledge of Microsoft 365',
                owner: 'Microsoft Learning',
                tags: ['Teams', 'Collaboration', 'Admin'],
                modules: [
                    { id: 'm1', title: 'Plan and Configure Microsoft Teams Environment', duration: '3h', status: 'not-started' },
                    { id: 'm2', title: 'Manage Chat, Teams, Channels, and Apps', duration: '4h', status: 'not-started' },
                    { id: 'm3', title: 'Manage Meetings and Calling', duration: '4h', status: 'not-started' },
                    { id: 'm4', title: 'Monitor and Troubleshoot Teams Environment', duration: '2h', status: 'not-started' }
                ],
                assessment: {
                    title: 'MS-700 Practice Assessment',
                    questions: 50,
                    passingScore: 70
                },
                description: 'Candidates for this exam are Microsoft Teams administrators who manage and configure Microsoft Teams in their organizations. They are responsible for configuring, deploying, and managing Office 365 workloads for Microsoft Teams that focus on efficient and effective collaboration and communication in an enterprise environment. Candidates must be able to plan, deploy, and manage Teams, chat, apps, channels, meetings, audio conferences, live events, and calling.',
                roles: ['Teams Administrator', 'IT Pro']
            },
            {
                id: 'tcta',
                name: 'Teams Calling Technical Assessment',
                code: 'TCTA',
                url: 'https://learn.microsoft.com/en-us/training/paths/manage-calling-options-microsoft-teams/',
                isMandatory: true,
                minRequired: 1,
                description: 'The Microsoft Teams Calling Technical Assessment verifies the ability of professionals to plan, configure, and manage calling features within the Teams environment. This includes subject matter expertise in managing PSTN connectivity through Direct Routing, Operator Connect, and Microsoft Teams Calling Plans. Professionals must ensure high-quality voice communications, configure emergency calling, and manage call policies and user settings to maintain a robust enterprise telephony solution.',
                roles: ['Voice Engineer', 'Teams Admin']
            }
        ]
    },
    {
        category: 'Custom Solutions for Microsoft Teams',
        level: 'Teams Apps Specialist',
        url: 'https://learn.microsoft.com/en-us/credentials/specializations/microsoft-teams/teams-apps-specialist',
        certs: [
            {
                id: 'ms-600',
                name: 'Microsoft 365 Developer Associate',
                code: 'MS-600',
                url: 'https://learn.microsoft.com/en-us/credentials/certifications/m365-developer-associate/',
                isMandatory: true,
                minRequired: 1,
                description: 'Microsoft 365 developers design, build, test, and maintain applications and solutions that are optimized for the productivity needs of organizations using the Microsoft 365 platform. They have deep knowledge of Microsoft Identity, Microsoft Graph, and Microsoft Teams. Candidates for this exam should be proficient in building apps for Microsoft Teams, extending SharePoint, and creating custom Office Add-ins using various coding techniques.',
                roles: ['Developer', 'App Builder']
            }
        ]
    },
    {
        category: 'Cloud Security',
        level: 'Security Specialization',
        url: 'https://learn.microsoft.com/en-us/credentials/specializations/security/cloud-security',
        certs: [
            {
                id: 'az-500',
                name: 'Microsoft Azure Security Technologies',
                code: 'AZ-500',
                url: 'https://learn.microsoft.com/en-us/credentials/certifications/azure-security-engineer/',
                isMandatory: true,
                minRequired: 1,
                description: 'Azure security engineers implement security controls and threat protection, manage identity and access, and protect data, applications, and networks in cloud and hybrid environments as part of an end-to-end infrastructure. This exam measures your ability to accomplish the following technical tasks: manage identity and access; implement platform protection; manage security operations; and secure data and applications in the Microsoft Azure ecosystem.',
                roles: ['Security Engineer', 'Cloud Admin']
            },
            {
                id: 'sc-200',
                name: 'Microsoft Security Operations Analyst',
                code: 'SC-200',
                url: 'https://learn.microsoft.com/en-us/credentials/certifications/security-operations-analyst/',
                isMandatory: true,
                minRequired: 1,
                description: 'Microsoft Security Operations Analysts collaborate with organizational stakeholders to secure information technology systems for the organization. Their goal is to reduce organizational risk by rapidly remediating active attacks in the environment, advising on improvements to threat protection practices, and referring violations of organizational policies to appropriate stakeholders. Responsibilities include threat management, monitoring, and response by using a variety of security solutions across their environment.',
                roles: ['Security Analyst', 'SOC Analyst']
            }
        ]
    },
    {
        category: 'Identity and Access Management',
        level: 'Security Specialization',
        url: 'https://learn.microsoft.com/en-us/credentials/specializations/security/identity-and-access-management',
        certs: [
            {
                id: 'sc-300',
                name: 'Identity and Access Administrator Associate',
                code: 'SC-300',
                url: 'https://learn.microsoft.com/en-us/credentials/certifications/identity-and-access-administrator/',
                isMandatory: true,
                minRequired: 1,
                description: 'The Microsoft Identity and Access Administrator designs, implements, and operates an organization’s identity and access management systems by using Azure AD. They provide seamless experiences and self-service management capabilities with adaptive access and governance of all identities. This role is crucial for securing modern applications and ensuring that only authorized users have access to sensitive corporate resources across diverse environments.',
                roles: ['IAM Specialist', 'Identity Admin']
            }
        ]
    },
    {
        category: 'Information Protection and Governance',
        level: 'Security Specialization',
        url: 'https://learn.microsoft.com/en-us/credentials/specializations/security/information-protection-and-governance',
        certs: [
            {
                id: 'sc-400',
                name: 'Information Protection and Compliance Administrator Associate',
                code: 'SC-400',
                url: 'https://learn.microsoft.com/en-us/credentials/certifications/information-protection-and-compliance-administrator/',
                isMandatory: true,
                minRequired: 1,
                description: 'Candidates for the SC-400 exam are responsible for planning and implementing controls that meet organizational compliance needs. They translate requirements and compliance controls into technical implementations and assist organizational control owners in becoming and staying compliant. They work with stakeholders to implement technology that supports data lifecycle management and data protection, including data loss prevention (DLP) and information protection policies.',
                roles: ['Compliance Administrator', 'Security Engineer', 'Data Protection Officer']
            }
        ]
    },
    {
        category: 'Threat Protection',
        level: 'Security Specialization',
        url: 'https://learn.microsoft.com/en-us/credentials/specializations/security/threat-protection',
        certs: [
            {
                id: 'sc-200-tp',
                name: 'Security Operations Analyst Associate',
                code: 'SC-200',
                url: 'https://learn.microsoft.com/en-us/credentials/certifications/security-operations-analyst/',
                isMandatory: true,
                minRequired: 1,
                description: 'The Security Operations Analyst Associate (SC-200) focuses on threat detection and response in Microsoft environments. They utilize Microsoft Sentinel, Microsoft Defender for Cloud, and Microsoft 365 Defender to hunt for and respond to security threats across the infrastructure. This role involves configuring data connectors, creating analytic rules, and managing incidents to ensure rapid recovery and robust security posture for the entire organization.',
                roles: ['Security Operations Analyst']
            }
        ]
    },
    {
        category: 'Adoption and Change Management',
        level: 'Modern Work Specialization',
        url: 'https://learn.microsoft.com/en-us/credentials/specializations/modern-work/adoption-and-change-management',
        certs: [
            {
                id: 'mw-assa',
                name: 'Adoption Service Specialist Assessment',
                code: 'MW-ASSA',
                url: 'https://learn.microsoft.com/en-us/training/paths/microsoft-service-adoption-specialist/',
                isMandatory: true,
                minRequired: 1,
                description: 'The Modern Work Adoption Service Specialist Assessment is designed for professionals who drive the utilization and adoption of Microsoft 365 services within an organization. This includes planning, launching, and managing adoption programs to ensure that employees are effectively using the tools provided to them. Professionals in this role focus on change management, stakeholder engagement, and measuring the success of adoption initiatives to maximize the ROI of Microsoft 365 investments.',
                roles: ['Adoption Specialist', 'Success Manager', 'Project Management', 'Project Manager']
            }
        ]
    },
    {
        category: 'Data Warehouse Migration to Microsoft Azure',
        level: 'Azure Specialization',
        url: 'https://learn.microsoft.com/en-us/credentials/specializations/azure/data-warehouse-migration',
        certs: [
            {
                id: 'dp-203',
                name: 'Azure Data Engineer Associate',
                code: 'DP-203',
                url: 'https://learn.microsoft.com/en-us/credentials/certifications/azure-data-engineer/',
                isMandatory: true,
                minRequired: 1,
                description: 'Microsoft Azure Data Engineers help stakeholders understand the data through exploration, and they build and maintain secure and compliant data processing pipelines by using different tools and techniques. These professionals use various Azure data services and languages to store and produce cleansed and enhanced datasets for analysis. This role involves designing and implementing data storage, data processing, and data security to support enterprise analytics requirements.',
                roles: ['Data Engineer', 'Analytics Engineer']
            },
            {
                id: 'dp-600',
                name: 'Fabric Analytics Engineer Associate',
                code: 'DP-600',
                url: 'https://learn.microsoft.com/en-us/credentials/certifications/fabric-analytics-engineer-associate/',
                isMandatory: true,
                minRequired: 1,
                description: 'The Fabric Analytics Engineer Associate (DP-600) is responsible for designing and deploying data analytics solutions at scale using Microsoft Fabric. This includes creating data models, transforming data into actionable insights, and managing data lifecycles within the Power BI and Fabric ecosystems. These engineers work closely with data scientists and business analysts to ensure that data architectures support advanced analytics and business intelligence requirements for the modern enterprise.',
                roles: ['Analytics Engineer', 'Data Engineer']
            }
        ]
    },
    {
        category: 'Accelerate Developer Productivity with Microsoft Azure',
        level: 'Azure Specialization',
        url: 'https://learn.microsoft.com/en-us/training/azure/',
        certs: [
            {
                id: 'az-400',
                name: 'DevOps Engineer Expert',
                code: 'AZ-400',
                url: 'https://learn.microsoft.com/en-us/credentials/certifications/devops-engineer/',
                isMandatory: true,
                minRequired: 1,
                description: 'DevOps engineers are developers or infrastructure administrators who also have subject matter expertise in working with people, processes, and products to enable continuous value delivery in organizations. They bridge the gap between development and operations teams, automating workflows, managing source code, and ensuring that software builds and deployments are efficient, secure, and reliable across various environments and platforms.',
                roles: ['DevOps Engineer', 'SRE', 'Project Management']
            }
        ]
    },
    {
        category: 'Hybrid Cloud Infrastructure with Microsoft Azure Stack HCI',
        level: 'Azure Specialization',
        url: 'https://learn.microsoft.com/en-us/training/paths/azure-stack-hci-foundations/',
        certs: [
            {
                id: 'az-104',
                name: 'Azure Administrator Associate',
                code: 'AZ-104',
                url: 'https://learn.microsoft.com/en-us/credentials/certifications/azure-administrator/',
                isMandatory: true,
                minRequired: 1,
                description: 'Candidates for this exam should have subject matter expertise in implementing, managing, and monitoring an organization’s Microsoft Azure environment. Responsibilities for this role include implementing, managing, and monitoring identity, governance, storage, compute, and virtual networks in a cloud environment, plus provision, size, monitor, and adjust resources, when needed. An Azure administrator often serves as part of a larger team dedicated to implementing an organization’s cloud infrastructure.',
                roles: ['Administrator', 'Cloud Engineer']
            },
            {
                id: 'az-305',
                name: 'Azure Solutions Architect Expert',
                code: 'AZ-305',
                url: 'https://learn.microsoft.com/en-us/credentials/certifications/azure-solutions-architect/',
                isMandatory: true,
                minRequired: 1,
                description: 'Azure solutions architects have subject matter expertise in designing cloud and hybrid solutions that run on Azure, including compute, network, storage, monitoring, and security. They collaborate with various stakeholders to understand requirement, design scalable and reliable solutions, and ensure that architectural best practices are followed. This role requires knowledge of virtualization, disaster recovery, networking, and business continuity.',
                roles: ['Solutions Architect', 'Cloud Architect', 'Project Manager']
            }
        ]
    },
    {
        category: 'Infra and Database Migration to Microsoft Azure',
        level: 'Azure Specialization',
        url: 'https://learn.microsoft.com/en-us/training/paths/migrate-sql-workloads-azure/',
        certs: [
            {
                id: 'az-400-db',
                name: 'DevOps Engineer Expert',
                code: 'AZ-400',
                url: 'https://learn.microsoft.com/en-us/credentials/certifications/devops-engineer/',
                isMandatory: true,
                minRequired: 1,
                description: 'DevOps engineers (AZ-400) bridge the gap between development and operations teams to enable continuous value delivery for organizations. They design and implement strategies for collaboration, code, infrastructure, source control, security, compliance, continuous integration, testing, delivery, deployment, and monitoring. This role is essential for maintaining robust and scalable enterprise software delivery pipelines.',
                roles: ['DevOps Engineer']
            },
            {
                id: 'az-104-db',
                name: 'Azure Administrator Associate',
                code: 'AZ-104',
                url: 'https://learn.microsoft.com/en-us/credentials/certifications/azure-administrator/',
                isMandatory: true,
                minRequired: 1,
                description: 'Candidates for the Azure Administrator Associate (AZ-104) certification should have subject matter expertise in implementing, managing, and monitoring an organization’s Microsoft Azure environment. This includes managing identity, governance, storage, compute, and virtual networks in a cloud environment. Administrators are responsible for monitoring and adjusting resources as needed, serving as a critical part of the larger IT team dedicated to cloud services.',
                roles: ['Administrator']
            },
            {
                id: 'az-500-db',
                name: 'Azure Security Engineer Associate',
                code: 'AZ-500',
                url: 'https://learn.microsoft.com/en-us/credentials/certifications/azure-security-engineer/',
                isMandatory: true,
                minRequired: 1,
                description: 'Azure security engineers (AZ-500) manage the security posture, identify and remediate vulnerabilities, and perform threat modeling within Azure environments. They implement threat protection and respond to security incident escalations. Security engineers often serve as part of a larger team dedicated to cloud-based management and security, ensuring that infrastructure remains compliant and protected against modern cyber threats.',
                roles: ['Security Engineer']
            },
            {
                id: 'dp-203-db',
                name: 'Azure Data Engineer Associate',
                code: 'DP-203',
                url: 'https://learn.microsoft.com/en-us/credentials/certifications/azure-data-engineer/',
                isMandatory: true,
                minRequired: 1,
                description: 'The Azure Data Engineer Associate (DP-203) integrates, transforms, and consolidates data from various structured and unstructured data systems into structures that are suitable for building analytics solutions. These professionals have subject matter expertise in using Azure data services to store and produce cleansed and enhanced datasets for analysis. They are responsible for designing and implementing data storage, security, and processing pipelines.',
                roles: ['Data Engineer']
            }
        ]
    },
    {
        category: 'AI Platform on Microsoft Azure',
        level: 'Azure Multidiscipline',
        url: 'https://learn.microsoft.com/en-us/training/paths/get-started-with-ai-on-azure/',
        certs: [
            {
                id: 'az-500-ai',
                name: 'Azure Security Engineer Associate',
                code: 'AZ-500',
                url: 'https://learn.microsoft.com/en-us/credentials/certifications/azure-security-engineer/',
                isMandatory: true,
                minRequired: 1,
                description: 'Security engineers for AI (AZ-500) implement security controls and threat protection specifically for artificial intelligence platforms. This role involves managing identity and access, protecting data and applications, and securing the entire AI lifecycle across cloud and hybrid environments. They ensure that AI models and datasets are protected from unauthorized access while maintaining operational efficiency for data science teams.',
                roles: ['Security Engineer', 'AI Specialist']
            },
            {
                id: 'dp-203-ai',
                name: 'Azure Data Engineer Associate',
                code: 'DP-203',
                url: 'https://learn.microsoft.com/en-us/credentials/certifications/azure-data-engineer/',
                isMandatory: true,
                minRequired: 1,
                description: 'Data engineers for AI (DP-203) design and implement data architectures that support advanced artificial intelligence and machine learning solutions. They develop data processing pipelines that feed high-quality data into AI models, ensuring scalability and reliability of the data foundation. This role is crucial for organizations looking to leverage AI for predictive analytics, automation, and intelligent decision-making.',
                roles: ['Data Engineer', 'AI Engineer']
            }
        ]
    },
    {
        category: 'Kubernetes on Microsoft Azure',
        level: 'App Innovation / Dev',
        url: 'https://learn.microsoft.com/en-us/training/paths/intro-to-kubernetes-on-azure/',
        certs: [
            {
                id: 'az-400-k8s',
                name: 'DevOps Engineer Expert',
                code: 'AZ-400',
                url: 'https://learn.microsoft.com/en-us/credentials/certifications/devops-engineer/',
                isMandatory: true,
                minRequired: 1,
                description: 'DevOps and Kubernetes orchestration.',
                roles: ['DevOps Engineer', 'Cloud Architect']
            },
            {
                id: 'az-204-k8s',
                name: 'Azure Developer Associate',
                code: 'AZ-204',
                url: 'https://learn.microsoft.com/en-us/credentials/certifications/azure-developer/',
                isMandatory: true,
                minRequired: 1,
                description: 'Developing solutions for Azure Kubernetes Service.',
                roles: ['Developer', 'Cloud Developer']
            },
            {
                id: 'az-104-k8s',
                name: 'Azure Administrator Associate',
                code: 'AZ-104',
                url: 'https://learn.microsoft.com/en-us/credentials/certifications/azure-administrator/',
                isMandatory: true,
                minRequired: 1,
                description: 'Cloud administration for Kubernetes environments.',
                roles: ['Administrator', 'Cloud Admin']
            }
        ]
    },
    {
        category: 'Microsoft Azure Virtual Desktop',
        level: 'Azure Specialization',
        url: 'https://learn.microsoft.com/en-us/credentials/specializations/azure/azure-virtual-desktop',
        certs: [
            {
                id: 'az-500-avd',
                name: 'Azure Security Engineer Associate',
                code: 'AZ-500',
                url: 'https://learn.microsoft.com/en-us/credentials/certifications/azure-security-engineer/',
                isMandatory: true,
                minRequired: 1,
                description: 'Security for virtual desktop infrastructure.',
                roles: ['Security Engineer']
            },
            {
                id: 'az-305-avd',
                name: 'Azure Solutions Architect Expert',
                code: 'AZ-305',
                url: 'https://learn.microsoft.com/en-us/credentials/certifications/azure-solutions-architect/',
                isMandatory: true,
                minRequired: 1,
                description: 'Azure solutions architects (AZ-305) have subject matter expertise in designing cloud and hybrid solutions that run on Azure, including compute, network, storage, monitoring, and security. They collaborate with business stakeholders and IT teams to design scalable, secure, and reliable solutions that align with the Microsoft Cloud Adoption Framework. This role requires knowledge of networking, virtualization, and business continuity strategies.',
                roles: ['Solutions Architect']
            },
            {
                id: 'az-104-avd',
                name: 'Azure Administrator Associate',
                code: 'AZ-104',
                url: 'https://learn.microsoft.com/en-us/credentials/certifications/azure-administrator/',
                isMandatory: true,
                minRequired: 1,
                description: 'Infrastructure management for AVD.',
                roles: ['Administrator']
            },
            {
                id: 'az-140-avd',
                name: 'Azure Virtual Desktop Specialty',
                code: 'AZ-140',
                url: 'https://learn.microsoft.com/en-us/credentials/certifications/azure-virtual-desktop-specialty/',
                isMandatory: true,
                minRequired: 1,
                description: 'Candidates for this certification are Azure administrators with subject matter expertise in planning, delivering, and managing virtual desktop experiences and remote apps.',
                roles: ['AVD Specialist', 'Cloud Admin']
            }
        ]
    },
    {
        category: 'Microsoft Azure VMware Solution',
        level: 'Azure Specialization',
        url: 'https://learn.microsoft.com/en-us/training/paths/azure-vmware-solution-foundations/',
        certs: [
            {
                id: 'az-104-avs',
                name: 'Azure Administrator Associate',
                code: 'AZ-104',
                url: 'https://learn.microsoft.com/en-us/credentials/certifications/azure-administrator/',
                isMandatory: true,
                minRequired: 1,
                description: 'Cloud administration for AVS.',
                roles: ['Administrator']
            },
            {
                id: 'az-305-avs',
                name: 'Azure Solutions Architect Expert',
                code: 'AZ-305',
                url: 'https://learn.microsoft.com/en-us/credentials/certifications/azure-solutions-architect/',
                isMandatory: true,
                minRequired: 1,
                description: 'The Azure Solutions Architect for AVS (AZ-305) designs hybrid VMware solutions that integrate on-premises VMware environments with the Azure VMware Solution. They are responsible for designing network connectivity, data migration strategies, and disaster recovery plans that bridge the gap between legacy virtualization and modern cloud platforms. This role ensures that VMware workloads can take full advantage of Azure’s scale and services.',
                roles: ['Solutions Architect']
            },
            {
                id: 'avs-tech',
                name: 'AVS Technical Assessment',
                code: 'AVS-TECH',
                url: 'https://learn.microsoft.com/en-us/training/paths/azure-vmware-solution-technical-assessment/',
                isMandatory: true,
                minRequired: 1,
                description: 'Technical assessment for Azure VMware Solution planning and implementation.',
                roles: ['Cloud Engineer', 'VMware Specialist']
            }
        ]
    },
    {
        category: 'Business Intelligence',
        level: 'Azure Specialization',
        url: 'https://learn.microsoft.com/en-us/training/paths/data-analysis-microsoft/',
        certs: [
            {
                id: 'pl-300-bi',
                name: 'Power BI Data Analyst Associate',
                code: 'PL-300',
                url: 'https://learn.microsoft.com/en-us/credentials/certifications/data-analyst-associate/',
                isMandatory: true,
                minRequired: 1,
                description: 'Microsoft Power BI data analysts deliver actionable insights by leveraging available data and applying domain expertise. They are responsible for cleaning, transforming, and modeling data using Power BI to create meaningful visualizations and reports. These professionals collaborate with business stakeholders to identify key performance indicators (KPIs) and provide data-driven recommendations to improve business processes and decision-making.',
                roles: ['Data Analyst', 'BI Developer']
            },
            {
                id: 'az-500-bi',
                name: 'Azure Security Engineer Associate',
                code: 'AZ-500',
                url: 'https://learn.microsoft.com/en-us/credentials/certifications/azure-security-engineer/',
                isMandatory: true,
                minRequired: 1,
                description: 'Security controls and threat protection for BI platforms.',
                roles: ['Security Engineer']
            },
            {
                id: 'dp-600-bi',
                name: 'Fabric Analytics Engineer Associate',
                code: 'DP-600',
                url: 'https://learn.microsoft.com/en-us/credentials/certifications/fabric-analytics-engineer-associate/',
                isMandatory: true,
                minRequired: 1,
                description: 'Designing, creating, and deploying enterprise-scale data analytics solutions.',
                roles: ['Analytics Engineer']
            }
        ]
    },
    {
        category: 'Finance',
        level: 'Biz Apps Specialization',
        url: 'https://learn.microsoft.com/en-us/training/paths/get-started-dynamics-365-finance/',
        certs: [
            {
                id: 'mb-310-fin',
                name: 'Dynamics 365 Finance Functional Consultant Associate',
                code: 'MB-310',
                url: 'https://learn.microsoft.com/en-us/credentials/certifications/dynamics-365-finance-functional-consultant-associate/',
                isMandatory: true,
                minRequired: 1,
                description: 'Microsoft Dynamics 365 Finance functional consultants unify global finances and operations; empower people to make fast, informed decisions; and help organizations adapt to changing market demands. They implement and configure features to meet business requirements, including financials, tax, budgeting, and asset management. These consultants work with various stakeholders to optimize financial performance and ensure compliance with regulatory standards and financial best practices.',
                roles: ['Functional Consultant', 'Finance Analyst', 'Finance Project Manager', 'Finance']
            },
            {
                id: 'mb-500-fin',
                name: 'Dynamics 365: Finance and Operations Apps Developer Associate',
                code: 'MB-500',
                url: 'https://learn.microsoft.com/en-us/credentials/certifications/dynamics-365-finance-and-operations-apps-developer-associate/',
                isMandatory: true,
                minRequired: 1,
                description: 'Designing, developing, securing, and extending Dynamics 365 Finance and Operations apps.',
                roles: ['Developer']
            },
            {
                id: 'mb-700-fin',
                name: 'Dynamics 365: Finance and Operations Apps Solution Architect Expert',
                code: 'MB-700',
                url: 'https://learn.microsoft.com/en-us/credentials/certifications/dynamics-365-finance-and-operations-apps-solution-architect/',
                isMandatory: true,
                minRequired: 1,
                description: 'Leading implementations and designing solutions.',
                roles: ['Solution Architect']
            }
        ]
    },
    {
        category: 'Intelligent Automation',
        level: 'Biz Apps Specialization',
        url: 'https://learn.microsoft.com/en-us/training/paths/automate-process-using-power-automate/',
        certs: [
            {
                id: 'pl-200-ia',
                name: 'Power Platform Functional Consultant Associate',
                code: 'PL-200',
                url: 'https://learn.microsoft.com/en-us/credentials/certifications/power-platform-functional-consultant-associate/',
                isMandatory: true,
                minRequired: 1,
                description: 'The Power Platform Functional Consultant Associate (PL-200) helps stakeholders solve business problems through the implementation of Microsoft Power Platform. They configure and implement solutions that include custom apps, automated workflows, and data visualizations. This role bridged the gap between business requirements and technical implementation, ensuring that low-code solutions deliver measurable value to the organization.',
                roles: ['Functional Consultant']
            },
            {
                id: 'pl-400-ia',
                name: 'Power Platform Developer Associate',
                code: 'PL-400',
                url: 'https://learn.microsoft.com/en-us/credentials/certifications/power-platform-developer-associate/',
                isMandatory: true,
                minRequired: 1,
                description: 'Power Platform Developers (PL-400) design, develop, secure, and troubleshoot Microsoft Power Platform solutions. They have subject matter expertise in using low-code tools as well as custom code (JavaScript, C#) to extend and customize the platform. This role involves building complex application logic, creating custom connectors, and integrating Power Platform with other external systems and cloud services.',
                roles: ['Developer']
            },
            {
                id: 'pl-500-ia',
                name: 'Power Automate RPA Developer Associate',
                code: 'PL-500',
                url: 'https://learn.microsoft.com/en-us/credentials/certifications/power-automate-rpa-developer-associate/',
                isMandatory: true,
                minRequired: 1,
                description: 'The Power Automate RPA Developer Associate (PL-500) automates repetitive, time-consuming business processes using Microsoft Power Automate and Power Automate Desktop. They analyze business requirements, design automation workflows, and implement unattended or attended RPA solutions. This role is essential for streamlining operations, reducing human error, and increasing efficiency across organizational workflows.',
                roles: ['RPA Developer', 'Automation Specialist', 'Project Management']
            },
            {
                id: 'pl-600-ia',
                name: 'Power Platform Solution Architect Expert',
                code: 'PL-600',
                url: 'https://learn.microsoft.com/en-us/credentials/certifications/power-platform-solution-architect/',
                isMandatory: true,
                minRequired: 1,
                description: 'Designing and implementing enterprise-scale solutions.',
                roles: ['Solution Architect']
            }
        ]
    },
    {
        category: 'Microsoft Low Code Application Development',
        level: 'Biz Apps Specialization',
        url: 'https://learn.microsoft.com/en-us/training/paths/build-low-code-apps-microsoft-power-platform/',
        certs: [
            {
                id: 'pl-200-lc',
                name: 'Power Platform Functional Consultant Associate',
                code: 'PL-200',
                url: 'https://learn.microsoft.com/en-us/credentials/certifications/power-platform-functional-consultant-associate/',
                isMandatory: true,
                minRequired: 1,
                description: 'Driving business value through low-code solutions.',
                roles: ['Functional Consultant']
            },
            {
                id: 'pl-400-lc',
                name: 'Power Platform Developer Associate',
                code: 'PL-400',
                url: 'https://learn.microsoft.com/en-us/credentials/certifications/power-platform-developer-associate/',
                isMandatory: true,
                minRequired: 1,
                description: 'The Power Platform Developer Associate (PL-400) creates custom apps and automated workflows using Microsoft Power Platform. They have subject matter expertise in build application components, including application enhancements, custom user experiences, system integrations, and data conversions. These developers collaborate with solution architects and business analysts to deliver robust and scalable low-code solutions across the enterprise.',
                roles: ['Developer']
            },
            {
                id: 'pl-600-lc',
                name: 'Power Platform Solution Architect Expert',
                code: 'PL-600',
                url: 'https://learn.microsoft.com/en-us/credentials/certifications/power-platform-solution-architect/',
                isMandatory: true,
                minRequired: 1,
                description: 'The Power Platform Solution Architect Expert (PL-600) is responsible for designing and implementing enterprise-scale solutions using the Power Platform ecosystem. This role requires deep technical knowledge of Power Apps, Power Automate, and Dataverse to orchestrate complex business transformations. Architects collaborate with stakeholders to ensure that solutions are secure, scalable, and provide maximum business value while adhering to architectural best practices.',
                roles: ['Solution Architect']
            }
        ]
    },
    {
        category: 'Sales',
        level: 'Biz Apps Specialization',
        url: 'https://learn.microsoft.com/en-us/training/paths/get-started-dynamics-365-sales/',
        certs: [
            {
                id: 'mb-210-sales',
                name: 'Dynamics 365 Sales Functional Consultant Associate',
                code: 'MB-210',
                url: 'https://learn.microsoft.com/en-us/credentials/certifications/dynamics-365-sales-functional-consultant-associate/',
                isMandatory: true,
                minRequired: 1,
                description: 'Microsoft Dynamics 365 Sales functional consultants are responsible for implementing solutions that maximize sales team productivity and enhance the customer relationship management experience. They configure sales features, manage lead and opportunity pipelines, and integrate with other Microsoft services like LinkedIn and Outlook. These consultants work to streamline sales processes, provide insights through dashboards, and ensure that sales teams have the tools they need to close deals effectively.',
                roles: ['Functional Consultant']
            },
            {
                id: 'mb-220-sales',
                name: 'Dynamics 365 Customer Insights (Journeys) Func Consultant',
                code: 'MB-220',
                url: 'https://learn.microsoft.com/en-us/credentials/certifications/dynamics-365-marketing-functional-consultant-associate/',
                isMandatory: true,
                minRequired: 1,
                description: 'Dynamics 365 Customer Insights (Journeys) functional consultants configure the marketing application to elevate the customer experience. They are responsible for creating customer segments, designing multi-channel marketing journeys, and optimizing marketing outcomes through data-driven insights. This role focuses on personalization, ensuring that marketing messages are delivered to the right audience at the right time across diverse channels.',
                roles: ['Functional Consultant']
            },
            {
                id: 'pl-600-sales',
                name: 'Power Platform Solution Architect Expert',
                code: 'PL-600',
                url: 'https://learn.microsoft.com/en-us/credentials/certifications/power-platform-solution-architect/',
                isMandatory: true,
                minRequired: 1,
                description: 'The Power Platform Solution Architect Expert (PL-600) is responsible for designing and implement enterprise-scale solutions using the Power Platform ecosystem. This role requires deep technical knowledge of Power Apps, Power Automate, and Dataverse to orchestrate complex business transformations. Architects collaborate with stakeholders to ensure that solutions are secure, scalable, and provide maximum business value while adhering to architectural best practices.',
                roles: ['Solution Architect', 'Project Manager']
            }
        ]
    }
];

export const getCertsByCompany = (companyId: string) => {
    if (companyId === 'microsoft') {
        return microsoftCertifications;
    }
    return [];
};
