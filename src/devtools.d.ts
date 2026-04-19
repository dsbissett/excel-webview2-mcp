type CSSInJS = string & {_tag: 'CSS-in-JS'};
declare module '*.css.js' {
  const styles: CSSInJS;
  export default styles;
}
