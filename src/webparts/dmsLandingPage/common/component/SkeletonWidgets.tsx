import { Skeleton, SkeletonItem } from "@fluentui/react-components";
import * as React from 'react';


const SkeletonWidgets = () => {
    return (
        <>
            <Skeleton aria-label="Loading card">
                <SkeletonItem shape="rectangle" style={{ width: "100%", height: "180px" }} />
                {/* <SkeletonItem style={{ width: "70%", height: "18px", marginTop: 5 }} />
                <SkeletonItem style={{ width: "40%", height: "16px", marginTop: 5 }} /> */}
            </Skeleton>

        </>
    );
};

export default React.memo(SkeletonWidgets);
