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

export const optimizeRequests = (tags: TagInfo[], maxGap: number = 20, maxBatchSize: number = 100): RequestBlock[] => {
    if (!tags || tags.length === 0) return [];

    // 1. Sort by address
    tags.sort((a, b) => a.Address - b.Address);

    const blocks: RequestBlock[] = [];
    
    // Ensure inputs are valid
    const SAFE_MAX_GAP = Math.max(0, maxGap);
    const SAFE_BATCH_SIZE = Math.max(1, maxBatchSize);

    let currentBlock: RequestBlock = {
        StartAddress: tags[0].Address,
        Length: tags[0].Length,
        IncludedTags: [tags[0]]
    };

    for (let i = 1; i < tags.length; i++) {
        const tag = tags[i];
        
        const currentEnd = currentBlock.StartAddress + currentBlock.Length;
        const gap = tag.Address - currentEnd;
        
        // Calculate the end of the current tag relative to the block start
        const tagEndRelative = (tag.Address + tag.Length) - currentBlock.StartAddress;
        
        // We must ensure the block is large enough to cover both the existing range AND the new tag.
        // If the new tag is "nested" inside the existing range, the length should not shrink.
        const newTotalLength = Math.max(currentBlock.Length, tagEndRelative);

        // Merge if gap is small and total size is within limit
        if (gap <= SAFE_MAX_GAP && newTotalLength <= SAFE_BATCH_SIZE) {
            currentBlock.Length = newTotalLength;
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