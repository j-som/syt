# import translateapi

class Solution:
    def findMedianSortedArrays(self, nums1, nums2) -> float:
        l1 = len(nums1)
        if (l1 == 0):
            return self.findMedianSortedArrarys1(nums2)
        l2 = len(nums2)
        if (l2 == 0):
            return self.findMedianSortedArrarys1(nums1)
        if (l1 == l2 and l1 == 1):
            return (nums1[0] + nums2[0]) / 2
        
        m1 = (l1 - 1) // 2
        a = nums1[m1]
        i12 = self.findSortedIndex(nums2, a)
        m2 = (l2 - 1) // 2
        b = nums2[m2]
        i21 = self.findSortedIndex(nums1, b)
#       nums1[m1]左边和nums2[i12]右边删掉同样数量的元素
#       nums2[m2]左边和nums1[i21]右边删掉同样数量的元素
        n1 = min(m1, l2 - i12)
        n2 = min(m2, l1 - i21)
        e1 = (l1 - 1 - n2)
        e2 = (l2 - 1 - n1)
        return self.findMedianSortedArrays(nums1[n1:e1], nums2[n2:e2])
        
    def findSortedIndex(self, nums, x: int) -> int:
        s = 0
        e = len(nums) - 1
        if (x > nums[e]): return e
        if (x <= nums[s]): return s
        while(e - s > 1):
            mid = (s + e) // 2
            if(nums[mid] >= x):
                e = mid
            else:
                s = mid
        
        if (nums[s] >= x): return s
        return e
    
    def findMedianSortedArrarys1(self, nums) -> float:
        start = 0
        end = len(nums) - 1
        mid = (start + end) // 2
        if (start + end) % 2 == 0:
            return nums[mid]
        else:
            return (nums[mid] + nums[mid+1]) / 2


if __name__ == '__main__':
    print(Solution().findMedianSortedArrays([1,3],[2]))