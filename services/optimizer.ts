export interface TagInfo {
    Name: string; // Full property name (e.g., Prop_M1)
    Address: number;
    Length: number; // Register count
    Type: string;
    ArrayLen: number; // Element count for arrays
}

export interface RequestBlock {
    StartAddress: number;
    Length: number; // Register count
    IncludedTags: TagInfo[];
}

export const optimizeRequests = (tags: TagInfo[]): RequestBlock[] => {
    if (!tags || tags.length === 0) return [];

    // 1. Sort by address
    tags.sort((a, b) => a.Address - b.Address);

    const blocks: RequestBlock[] = [];
    const MAX_GAP = 20;
    const MAX_BATCH_SIZE = 100;

    let currentBlock: RequestBlock = {
        StartAddress: tags[0].Address,
        Length: tags[0].Length,
        IncludedTags: [tags[0]]
    };

    for (let i = 1; i < tags.length; i++) {
        const tag = tags[i];
        
        const currentEnd = currentBlock.StartAddress + currentBlock.Length;
        const gap = tag.Address - currentEnd;
        const newLength = (tag.Address + tag.Length) - currentBlock.StartAddress;

        // Merge if gap is small and total size is within limit
        if (gap <= MAX_GAP && newLength <= MAX_BATCH_SIZE) {
            currentBlock.Length = newLength;
            currentBlock.IncludedTags.push(tag);
        } else {
            blocks.push(currentBlock);
            currentBlock = {
                StartAddress: tag.Address,
                Length: tag.Length,
                IncludedTags: [tag]
            };
        }
    }
    blocks.push(currentBlock);

    return blocks;
};
