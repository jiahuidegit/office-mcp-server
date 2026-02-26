/**
 * 主题工厂函数和工具
 */

import { Theme } from './types.js'

/** 厘米转 twips (1cm = 567 twips) */
export const cmToTwips = (cm: number) => Math.round(cm * 567)

/** 默认页面布局：上下2.54cm 左右3.17cm */
const defaultPageLayout = {
  marginTop: cmToTwips(2.54),
  marginBottom: cmToTwips(2.54),
  marginLeft: cmToTwips(3.17),
  marginRight: cmToTwips(3.17)
}

/**
 * 创建主题，合并默认值
 * 只需传入与默认值不同的部分
 */
export function createTheme(partial: Partial<Theme> & Pick<Theme, 'name' | 'displayName' | 'description' | 'colors' | 'word' | 'excel'>): Theme {
  return {
    pageLayout: defaultPageLayout,
    ...partial
  }
}
